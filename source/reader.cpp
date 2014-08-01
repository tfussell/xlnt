#include <algorithm>
#include <pugixml.hpp>

#include "reader/reader.hpp"
#include "cell/cell.hpp"
#include "cell/value.hpp"
#include "common/datetime.hpp"
#include "worksheet/range_reference.hpp"
#include "workbook/workbook.hpp"
#include "worksheet/worksheet.hpp"
#include "workbook/document_properties.hpp"
#include "common/relationship.hpp"
#include "common/zip_file.hpp"
#include "common/exceptions.hpp"

namespace xlnt {

const std::string reader::CentralDirectorySignature = "\x50\x4b\x05\x06";

std::string::size_type find_string_in_string(const std::string &string, const std::string &substring)
{
    std::string::size_type possible_match_index = string.find(substring.at(0));

    while(possible_match_index != std::string::npos)
    {
        if(string.substr(possible_match_index, substring.size()) == substring)
        {
            return possible_match_index;
        }

        possible_match_index = string.find(substring.at(0), possible_match_index + 1);
    }

    return possible_match_index;
}

std::string reader::repair_central_directory(const std::string &original)
{
    auto pos = find_string_in_string(original, CentralDirectorySignature);

    if(pos != std::string::npos)
    {
        return original.substr(0, pos + 22);
    }

    return original;
}

std::vector<std::pair<std::string, std::string>> reader::read_sheets(zip_file &archive)
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

datetime w3cdtf_to_datetime(const std::string &string)
{
    datetime result(1900, 1, 1);
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

document_properties reader::read_properties_core(const std::string &xml_string)
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

std::string reader::read_dimension(const std::string &xml_string)
{
    pugi::xml_document doc;
    doc.load(xml_string.c_str());
    auto root_node = doc.child("worksheet");
    auto dimension_node = root_node.child("dimension");
    std::string dimension = dimension_node.attribute("ref").as_string();
    return dimension;
}

std::vector<relationship> reader::read_relationships(zip_file &archive, const std::string &filename)
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

std::vector<std::pair<std::string, std::string>> reader::read_content_types(zip_file &archive)
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

std::string reader::determine_document_type(const std::vector<std::pair<std::string, std::string>> &override_types)
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

void read_worksheet_common(worksheet ws, const pugi::xml_node &root_node, const std::vector<std::string> &string_table, const std::vector<int> &number_format_ids)
{
    auto dimension_node = root_node.child("dimension");
    std::string dimension = dimension_node.attribute("ref").as_string();
    auto sheet_data_node = root_node.child("sheetData");
    auto merge_cells_node = root_node.child("mergeCells");

    if(merge_cells_node != nullptr)
    {
        int count = merge_cells_node.attribute("count").as_int();

        for(auto merge_cell_node : merge_cells_node.children("mergeCell"))
        {
            ws.merge_cells(merge_cell_node.attribute("ref").as_string());
            count--;
        }

        if(count != 0)
        {
            throw std::runtime_error("mismatch between count and actual number of merged cells");
        }
    }

    for(auto row_node : sheet_data_node.children("row"))
    {
        int row_index = row_node.attribute("r").as_int();
        std::string span_string = row_node.attribute("spans").as_string();
        auto colon_index = span_string.find(':');

        if(colon_index == std::string::npos)
        {
            continue;
        }

        int min_column = std::stoi(span_string.substr(0, colon_index));
        int max_column = std::stoi(span_string.substr(colon_index + 1));

        for(int i = min_column; i < max_column + 1; i++)
        {
            std::string address = xlnt::cell_reference::column_string_from_index(i) + std::to_string(row_index);
            auto cell_node = row_node.find_child_by_attribute("c", "r", address.c_str());

            if(cell_node != nullptr)
            {
                bool has_value = cell_node.child("v") != nullptr;
                std::string value_string = cell_node.child("v").text().as_string();

                bool has_type = cell_node.attribute("t") != nullptr;
                std::string type = cell_node.attribute("t").as_string();

                bool has_style = cell_node.attribute("s") != nullptr;
                std::string style = cell_node.attribute("s").as_string();

                bool has_formula = cell_node.child("f") != nullptr;
                bool shared_formula = has_formula && cell_node.child("f").attribute("t") != nullptr && std::string(cell_node.child("f").attribute("t").as_string()) == "shared";

                if(has_formula && !shared_formula && !ws.get_parent().get_data_only())
                {
                    std::string formula = cell_node.child("f").text().as_string();
                    ws.get_cell(address).set_formula(formula);
                }

                if(has_type && type == "inlineStr") // inline string
                {
                    std::string inline_string = cell_node.child("is").child("t").text().as_string();
                    ws.get_cell(address).set_value(inline_string);
                }
                else if(has_type && type == "s") // shared string
                {
                    auto shared_string_index = std::stoi(value_string);
                    auto shared_string = string_table.at(shared_string_index);
                    ws.get_cell(address).set_value(shared_string);
                }
                else if(has_type && type == "b") // boolean
                {
                    ws.get_cell(address).set_value(value(value_string != "0"));
                }
                else if(has_type && type == "str")
                {
                    ws.get_cell(address).set_value(value_string);
                }
                else if(has_style)
                {
                    auto number_format_id = number_format_ids.at(std::stoi(style));
                    auto format = number_format::lookup_format(number_format_id);
                    ws.get_cell(address).get_style().get_number_format().set_format_code(format);
                    if(format == number_format::format::date_xlsx14)
                    {
                        auto base_date = ws.get_parent().get_properties().excel_base_date;
                        auto converted = date::from_number(std::stoi(value_string), base_date);
                        ws.get_cell(address).set_value(converted.to_number(calendar::windows_1900));
                    }
                    else
                    {
                        ws.get_cell(address).set_value(value(std::stod(value_string)));
                    }
                }
                else if(has_value)
                {
                    try
                    {
                        ws.get_cell(address).set_value(value(std::stod(value_string)));
                    }
                    catch(std::invalid_argument)
                    {
                        ws.get_cell(address).set_value(value_string);
                    }
                }
            }
        }
    }

    auto auto_filter_node = root_node.child("autoFilter");

    if(auto_filter_node != nullptr)
    {
        range_reference ref(auto_filter_node.attribute("ref").as_string());
        ws.auto_filter(ref);
    }
}

void reader::fast_parse(worksheet ws, std::istream &xml_source, const std::vector<std::string> &shared_string, const std::vector<style> &/*style_table*/, std::size_t /*color_index*/)
{
    pugi::xml_document doc;
    doc.load(xml_source);
    read_worksheet_common(ws, doc.child("worksheet"), shared_string, {});
}

void reader::read_worksheet(worksheet ws, const std::string &xml_string, const std::vector<std::string> &string_table, const std::vector<int> &number_format_ids)
{
    pugi::xml_document doc;
    doc.load(xml_string.c_str());
    read_worksheet_common(ws, doc.child("worksheet"), string_table, number_format_ids);
}

worksheet xlnt::reader::read_worksheet(std::istream &handle, xlnt::workbook &wb, const std::string &title, const std::vector<std::string> &string_table)
{
    auto ws = wb.create_sheet();
    ws.set_title(title);
    pugi::xml_document doc;
    doc.load(handle);
    read_worksheet_common(ws, doc.child("worksheet"), string_table, {});
    return ws;
}

std::vector<std::string> reader::read_shared_string(const std::string &xml_string)
{
    std::vector<std::string> shared_strings;
    pugi::xml_document doc;
    doc.load(xml_string.c_str());
    auto root_node = doc.child("sst");
    //int count = root_node.attribute("count").as_int();
    int unique_count = root_node.attribute("uniqueCount").as_int();

    for(auto si_node : root_node)
    {
        shared_strings.push_back(si_node.child("t").text().as_string());
    }

    if(unique_count != (int)shared_strings.size())
    {
        throw std::runtime_error("counts don't match");
    }

    return shared_strings;
}

workbook reader::load_workbook(const std::string &filename, bool guess_types, bool data_only)
{
    workbook wb;
    wb.set_guess_types(guess_types);
    wb.set_data_only(data_only);
    wb.load(filename);
    return wb;
}

std::vector<std::pair<std::string, std::string>> reader::detect_worksheets(zip_file &archive)
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

    auto workbook_relationships = reader::read_relationships(archive, "xl/workbook.xml");
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

} // namespace xlnt
