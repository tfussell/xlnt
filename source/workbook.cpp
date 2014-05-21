#include <algorithm>
#include <sstream>
#include <pugixml.hpp>

#include "workbook.h"
#include "custom_exceptions.h"
#include "drawing.h"
#include "named_range.h"
#include "reader.h"
#include "relationship.h"
#include "worksheet.h"
#include "writer.h"
#include "zip_file.h"

namespace xlnt {

workbook::workbook(optimization optimization)
: optimized_write_(optimization==optimization::write),
optimized_read_(optimization==optimization::read),
active_sheet_index_(0)
{
    if(!optimized_write_)
    {
        create_sheet("Sheet1");
    }
}

workbook::~workbook()
{
    clear();
}

worksheet workbook::get_sheet_by_name(const std::string &name)
{
    auto match = std::find_if(worksheets_.begin(), worksheets_.end(), [&](const worksheet &w) { return w.get_title() == name; });
    if(match != worksheets_.end())
    {
        return worksheet(*match);
    }
    return worksheet(nullptr);
}

worksheet workbook::get_active_sheet()
{
    return worksheets_[active_sheet_index_];
}

bool workbook::has_named_range(const std::string &name, xlnt::worksheet ws) const
{
    for(auto named_range : named_ranges_)
    {
        if(named_range.first == name && named_range.second.get_scope() == ws)
        {
            return true;
        }
    }
    return false;
}

worksheet workbook::create_sheet()
{
    if(optimized_read_)
    {
        throw std::runtime_error("this is a read-only workbook");
    }
    
    std::string title = "Sheet1";
    int index = 1;
    while(get_sheet_by_name(title) != nullptr)
    {
        title = "Sheet" + std::to_string(++index);
    }
    
    worksheet ws = worksheet(worksheet::allocate(*this, title));
    worksheets_.push_back(ws);
    
    return get_sheet_by_name(title);
}

void workbook::add_sheet(xlnt::worksheet worksheet)
{
    if(optimized_read_)
    {
        throw std::runtime_error("this is a read-only workbook");
    }
    
    for(auto ws : *this)
    {
        if(worksheet == ws)
        {
            throw std::runtime_error("worksheet already in workbook");
        }
    }
    worksheets_.push_back(worksheet);
}

void workbook::add_sheet(xlnt::worksheet worksheet, std::size_t index)
{
    if(optimized_read_)
    {
        throw std::runtime_error("this is a read-only workbook");
    }
    
    for(auto ws : *this)
    {
        if(worksheet == ws)
        {
            throw std::runtime_error("worksheet already in workbook");
        }
    }
    worksheets_.push_back(worksheet);
    std::swap(worksheets_[index], worksheets_.back());
}

int workbook::get_index(xlnt::worksheet worksheet)
{
    int i = 0;
    for(auto ws : *this)
    {
        if(worksheet == ws)
        {
            return i;
        }
        i++;
    }
    throw std::runtime_error("worksheet isn't owned by this workbook");
}

named_range workbook::create_named_range(const std::string &name, worksheet range_owner, const range_reference &reference)
{
    for(auto ws : worksheets_)
    {
        if(ws == range_owner)
        {
            named_ranges_[name] = range_owner.create_named_range(name, reference);
            return named_ranges_[name];
        }
    }
    
    throw std::runtime_error("worksheet isn't owned by this workbook");
}

void workbook::add_named_range(const std::string &name, named_range named_range)
{
    for(auto ws : worksheets_)
    {
        if(ws == named_range.get_scope())
        {
            named_ranges_[name] = named_range;
            return;
        }
    }
    
    throw std::runtime_error("worksheet isn't owned by this workbook");
}

std::vector<named_range> workbook::get_named_ranges()
{
    std::vector<named_range> named_ranges;
    for(auto named_range : named_ranges_)
    {
        named_ranges.push_back(named_range.second);
    }
    return named_ranges;
}

void workbook::remove_named_range(named_range named_range)
{
    std::string key_match = "";
    for(auto r : named_ranges_)
    {
        if(r.second == named_range)
        {
            key_match = r.first;
        }
    }
    if(key_match == "")
    {
        throw std::runtime_error("range not found in worksheet");
    }
    named_ranges_.erase(key_match);
}

named_range workbook::get_named_range(const std::string &name, worksheet ws)
{
    for(auto named_range : named_ranges_)
    {
        if(named_range.first == name && named_range.second.get_scope() == ws)
        {
            return named_range.second;
        }
    }
    throw std::runtime_error("doesn't exist");
}

void workbook::load(const std::string &filename)
{
    zip_file f(filename, file_mode::open);
    //auto core_properties = read_core_properties();
    //auto app_properties = read_app_properties();
    auto root_relationships = reader::read_relationships(f.get_file_contents("_rels/.rels"));
    auto content_types = reader::read_content_types(f.get_file_contents("[Content_Types].xml"));
    
    auto type = reader::determine_document_type(root_relationships, content_types.second);
    
    if(type != "excel")
    {
        throw std::runtime_error("unsupported document type: " + filename);
    }
    
    auto workbook_relationships = reader::read_relationships(f.get_file_contents("xl/_rels/workbook.xml.rels"));
    
    pugi::xml_document doc;
    doc.load(f.get_file_contents("xl/workbook.xml").c_str());
    
    auto root_node = doc.child("workbook");
    auto sheets_node = root_node.child("sheets");
    
    while(!worksheets_.empty())
    {
        remove_sheet(worksheets_.front());
    }
    
    std::vector<std::string> shared_strings;
    if(f.has_file("xl/sharedStrings.xml"))
    {
        shared_strings = xlnt::reader::read_shared_string(f.get_file_contents("xl/sharedStrings.xml"));
    }
    
    for(auto sheet_node : sheets_node.children("sheet"))
    {
        auto relation_id = sheet_node.attribute("r:id").as_string();
        auto ws = create_sheet(sheet_node.attribute("name").as_string());
        std::string sheet_filename("xl/");
        sheet_filename += workbook_relationships[relation_id].second;
        xlnt::reader::read_worksheet(ws, f.get_file_contents(sheet_filename).c_str(), shared_strings);
    }
}

void workbook::remove_sheet(worksheet ws)
{
    auto match_iter = std::find(worksheets_.begin(), worksheets_.end(), ws);
    if(match_iter == worksheets_.end())
    {
        throw std::runtime_error("worksheet not owned by this workbook");
    }
    worksheet::deallocate(*match_iter);
    worksheets_.erase(match_iter);
}

worksheet workbook::create_sheet(std::size_t index)
{
    auto ws = create_sheet();
    if(index != worksheets_.size())
    {
        std::swap(worksheets_[index], worksheets_.back());
    }
    return ws;
}

worksheet workbook::create_sheet(std::size_t index, const std::string &title)
{
    auto ws = create_sheet(index);
    ws.set_title(title);
    return ws;
}

worksheet workbook::create_sheet(const std::string &title)
{
    if(title.length() > 31)
    {
        throw bad_sheet_title(title);
    }
    
    if(std::find_if(title.begin(), title.end(),
                    [](char c) { return c == '*' || c == ':' || c == '/' || c == '\\' || c == '?' || c == '[' || c == ']'; }) != title.end())
    {
        throw bad_sheet_title(title);
    }
    
    if(get_sheet_by_name(title) != nullptr)
    {
        throw std::runtime_error("sheet exists");
    }
    
    auto ws = worksheet::allocate(*this, title);
    worksheets_.push_back(ws);
    return worksheet(ws);
}

std::vector<worksheet>::iterator workbook::begin()
{
    return worksheets_.begin();
}

std::vector<worksheet>::iterator workbook::end()
{
    return worksheets_.end();
}

std::vector<std::string> workbook::get_sheet_names() const
{
    std::vector<std::string> names;
    for(auto &ws : worksheets_)
    {
        names.push_back(ws.get_title());
    }
    return names;
}

worksheet workbook::operator[](const std::string &name)
{
    for(auto sheet : worksheets_)
    {
        if(sheet.get_title() == name)
        {
            return sheet;
        }
    }
    throw std::runtime_error("sheet not found");
}

worksheet workbook::operator[](int index)
{
    return worksheets_[index];
}

void workbook::clear()
{
    while(!worksheets_.empty())
    {
        worksheet::deallocate(worksheets_.back());
        worksheets_.pop_back();
    }
}

void workbook::save(const std::string &filename)
{
    zip_file f(filename, file_mode::create, file_access::write);
    
    std::pair<std::unordered_map<std::string, std::string>, std::unordered_map<std::string, std::string>> content_types =
    {
        {
            {"rels", "application/vnd.openxmlformats-package.relationships+xml"},
            {"xml", "application/xml"}
        },
        {
            {"/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"},
            {"/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"},
            {"/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"},
            {"/xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"},
            {"/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"},
            {"/xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"}
        }
    };
    
    f.set_file_contents("[Content_Types].xml", writer::write_content_types(content_types));
    
    std::unordered_map<std::string, std::pair<std::string, std::string>> root_rels =
    {
        {"rId3", {"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml"}},
        {"rId2", {"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml"}},
        {"rId1", {"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml"}}
    };
    f.set_file_contents("_rels/.rels", writer::write_relationships(root_rels));
    
    std::unordered_map<std::string, std::pair<std::string, std::string>> workbook_rels =
    {
        {"rId1", {"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "worksheets/sheet1.xml"}},
        {"rId2", {"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml"}},
        {"rId3", {"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml"}}
    };
    f.set_file_contents("xl/_rels/workbook.xml.rels", writer::write_relationships(workbook_rels));
    
    pugi::xml_document doc;
    auto root_node = doc.append_child("workbook");
    root_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    root_node.append_attribute("xmlns:r").set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    auto sheets_node = root_node.append_child("sheets");
    
    int i = 0;
    for(auto ws : *this)
    {
        auto sheet_node = sheets_node.append_child("sheet");
        sheet_node.append_attribute("name").set_value(ws.get_title().c_str());
        sheet_node.append_attribute("sheetId").set_value(std::to_string(i + 1).c_str());
        sheet_node.append_attribute("r:id").set_value((std::string("rId") + std::to_string(i + 1)).c_str());
        std::string filename = "xl/worksheets/sheet";
        f.set_file_contents(filename + std::to_string(i + 1) + ".xml", xlnt::writer::write_worksheet(ws));
        i++;
    }
    
    std::stringstream ss;
    doc.save(ss);
    
    f.set_file_contents("xl/workbook.xml", ss.str());
}
    
bool workbook::operator==(const workbook &rhs) const
{
    if(optimized_write_ == rhs.optimized_write_
       && optimized_read_ == rhs.optimized_read_
       && guess_types_ == rhs.guess_types_
       && data_only_ == rhs.data_only_
       && active_sheet_index_ == rhs.active_sheet_index_)
    {
        if(worksheets_.size() != rhs.worksheets_.size())
        {
            return false;
        }
        
        for(std::size_t i = 0; i < worksheets_.size(); i++)
        {
            if(worksheets_[i] != rhs.worksheets_[i])
            {
                return false;
            }
        }
        
        /*
         if(named_ranges_.size() != rhs.named_ranges_.size())
         {
         return false;
         }
         
         for(int i = 0; i < named_ranges_.size(); i++)
         {
         if(named_ranges_[i] != rhs.named_ranges_[i])
         {
         return false;
         }
         }
         
         if(relationships_.size() != rhs.relationships_.size())
         {
         return false;
         }
         
         for(int i = 0; i < relationships_.size(); i++)
         {
         if(relationships_[i] != rhs.relationships_[i])
         {
         return false;
         }
         }
         
         if(drawings_.size() != rhs.drawings_.size())
         {
         return false;
         }
         
         for(int i = 0; i < drawings_.size(); i++)
         {
         if(drawings_[i] != rhs.drawings_[i])
         {
         return false;
         }
         }
         */
        
        return true;
    }
    return false;
}
    
}