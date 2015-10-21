#include <xlnt/common/datetime.hpp>
#include <xlnt/common/exceptions.hpp>
#include <xlnt/common/zip_file.hpp>
#include <xlnt/reader/workbook_reader.hpp>
#include <xlnt/workbook/document_properties.hpp>
#include <xlnt/workbook/workbook.hpp>

#include "detail/include_pugixml.hpp"

namespace {

xlnt::datetime w3cdtf_to_datetime(const std::string &string)
{
    xlnt::datetime result(1900, 1, 1);
    auto separator_index = string.find('-');
    result.year = std::stoi(string.substr(0, separator_index));
    result.month = std::stoi(string.substr(separator_index + 1, string.find('-', separator_index + 1)));
    separator_index = string.find('-', separator_index + 1);
    result.day = std::stoi(string.substr(separator_index + 1, string.find('T', separator_index + 1)));
    separator_index = string.find('T', separator_index + 1);
    result.hour = std::stoi(string.substr(separator_index + 1, string.find(':', separator_index + 1)));
    separator_index = string.find(':', separator_index + 1);
    result.minute = std::stoi(string.substr(separator_index + 1, string.find(':', separator_index + 1)));
    separator_index = string.find(':', separator_index + 1);
    result.second = std::stoi(string.substr(separator_index + 1, string.find('Z', separator_index + 1)));
    return result;
}
    
}

namespace xlnt {

std::vector<std::pair<std::string, std::string>> read_sheets(zip_file &archive)
{
    auto xml_source = archive.read("xl/workbook.xml");
    pugi::xml_document doc;
    doc.load(xml_source.c_str());
    
    std::string ns;
    
    for(auto child : doc.children())
    {
        std::string name = child.name();
        
        if(name.find(':') != std::string::npos)
        {
            auto colon_index = name.find(':');
            ns = name.substr(0, colon_index);
            break;
        }
    }
    
    auto with_ns = [&](const std::string &base) { return ns.empty() ? base : ns + ":" + base; };
    
    auto root_node = doc.child(with_ns("workbook").c_str());
    auto sheets_node = root_node.child(with_ns("sheets").c_str());
    
    std::vector<std::pair<std::string, std::string>> sheets;
    
    // store temp because pugixml iteration uses the internal char array multiple times
    auto sheet_element_name = with_ns("sheet");
    
    for(auto sheet_node : sheets_node.children(sheet_element_name.c_str()))
    {
        std::string id = sheet_node.attribute("r:id").as_string();
        std::string name = sheet_node.attribute("name").as_string();
        sheets.push_back(std::make_pair(id, name));
    }
    
    return sheets;
}

document_properties read_properties_core(const std::string &xml_string)
{
    document_properties props;
    pugi::xml_document doc;
    doc.load(xml_string.c_str());
    auto root_node = doc.child("cp:coreProperties");
    
    props.excel_base_date = calendar::windows_1900;
    
    if(root_node.child("dc:creator") != nullptr)
    {
        props.creator = root_node.child("dc:creator").text().as_string();
    }
    if(root_node.child("cp:lastModifiedBy") != nullptr)
    {
        props.last_modified_by = root_node.child("cp:lastModifiedBy").text().as_string();
    }
    if(root_node.child("dcterms:created") != nullptr)
    {
        std::string created_string = root_node.child("dcterms:created").text().as_string();
        props.created = w3cdtf_to_datetime(created_string);
    }
    if(root_node.child("dcterms:modified") != nullptr)
    {
        std::string modified_string = root_node.child("dcterms:modified").text().as_string();
        props.modified = w3cdtf_to_datetime(modified_string);
    }
    
    return props;
}

std::string read_dimension(const std::string &xml_string)
{
    pugi::xml_document doc;
    doc.load(xml_string.c_str());
    auto root_node = doc.child("worksheet");
    auto dimension_node = root_node.child("dimension");
    std::string dimension = dimension_node.attribute("ref").as_string();
    return dimension;
}

std::vector<relationship> read_relationships(zip_file &archive, const std::string &filename)
{
    auto filename_separator_index = filename.find_last_of('/');
    auto basename = filename.substr(filename_separator_index + 1);
    auto dirname = filename.substr(0, filename_separator_index);
    auto rels_filename = dirname + "/_rels/" + basename + ".rels";
    
    pugi::xml_document doc;
    auto content = archive.read(rels_filename);
    doc.load(content.c_str());
    
    auto root_node = doc.child("Relationships");
    
    std::vector<relationship> relationships;
    
    for(auto relationship : root_node.children("Relationship"))
    {
        std::string id = relationship.attribute("Id").as_string();
        std::string type = relationship.attribute("Type").as_string();
        std::string target = relationship.attribute("Target").as_string();
        
        if(target[0] != '/' && target.substr(0, 2) != "..")
        {
            target = dirname + "/" + target;
        }
        
        if(target[0] == '/')
        {
            target = target.substr(1);
        }
        
        relationships.push_back(xlnt::relationship(type, id, target));
    }
    
    return relationships;
}

std::vector<std::pair<std::string, std::string>> read_content_types(zip_file &archive)
{
    pugi::xml_document doc;
    
    try
    {
        auto content_types_string = archive.read("[Content_Types].xml");
        doc.load(content_types_string.c_str());
    }
    catch(std::exception e)
    {
        throw invalid_file_exception(archive.get_filename());
    }
    
    auto root_node = doc.child("Types");
    
    std::vector<std::pair<std::string, std::string>> override_types;
    
    for(auto child : root_node.children("Override"))
    {
        std::string part_name = child.attribute("PartName").as_string();
        std::string content_type = child.attribute("ContentType").as_string();
        override_types.push_back({part_name, content_type});
    }
    
    return override_types;
}

std::string determine_document_type(const std::vector<std::pair<std::string, std::string>> &override_types)
{
    auto match = std::find_if(override_types.begin(), override_types.end(), [](const std::pair<std::string, std::string> &p) { return p.first == "/xl/workbook.xml"; });
    
    if(match == override_types.end())
    {
        return "unsupported";
    }
    
    std::string type = match->second;
    
    if(type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    {
        return "excel";
    }
    else if(type == "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")
    {
        return "powerpoint";
    }
    else if(type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
    {
        return "word";
    }
    
    return "unsupported";
}

std::vector<std::pair<std::string, std::string>> detect_worksheets(zip_file &archive)
{
    static const std::string ValidWorksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    
    auto content_types = read_content_types(archive);
    std::vector<std::string> valid_sheets;
    
    for(const auto &content_type : content_types)
    {
        if(content_type.second == ValidWorksheet)
        {
            valid_sheets.push_back(content_type.first);
        }
    }
    
    auto workbook_relationships = read_relationships(archive, "xl/workbook.xml");
    std::vector<std::pair<std::string, std::string>> result;
    
    for(const auto &ws : read_sheets(archive))
    {
        auto rel = *std::find_if(workbook_relationships.begin(), workbook_relationships.end(), [&](const relationship &r) { return r.get_id() == ws.first; });
        auto target = rel.get_target_uri();
        
        if(std::find(valid_sheets.begin(), valid_sheets.end(), "/" + target) != valid_sheets.end())
        {
            result.push_back({target, ws.second});
        }
    }
    
    return result;
}

}
