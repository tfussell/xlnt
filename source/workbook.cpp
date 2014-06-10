#include <algorithm>
#include <fstream>
#include <sstream>
#include <pugixml.hpp>

#include "workbook/workbook.hpp"
#include "common/exceptions.hpp"
#include "drawing/drawing.hpp"
#include "worksheet/range.hpp"
#include "reader/reader.hpp"
#include "common/relationship.hpp"
#include "worksheet/worksheet.hpp"
#include "writer/writer.hpp"
#include "common/zip_file.hpp"
#include "detail/cell_impl.hpp"
#include "detail/workbook_impl.hpp"
#include "detail/worksheet_impl.hpp"

namespace xlnt {
namespace detail {
workbook_impl::workbook_impl(optimization o) : optimized_read_(o == optimization::read), optimized_write_(o == optimization::write), active_sheet_index_(0)
{
    
}
} // namespace detail
    
workbook::workbook(optimization optimize) : d_(new detail::workbook_impl(optimize))
{
    if(!d_->optimized_read_)
    {
        create_sheet("Sheet");
    }
}

workbook::~workbook()
{
    clear();
}

workbook::iterator::iterator(workbook &wb, std::size_t index) : wb_(wb), index_(index)
{
    
}

worksheet workbook::iterator::operator*()
{
    return wb_[index_];
}
    
workbook::iterator &workbook::iterator::operator++()
{
    index_++;
    return *this;
}
    
workbook::iterator workbook::iterator::operator++(int)
{
    iterator old(wb_, index_);
    ++*this;
    return old;
}
    
bool workbook::iterator::operator==(const iterator &comparand) const
{
    return index_ == comparand.index_ && wb_ == comparand.wb_;
}
    
workbook::const_iterator::const_iterator(const workbook &wb, std::size_t index) : wb_(wb), index_(index)
{
    
}

const worksheet workbook::const_iterator::operator*()
{
    return wb_.get_sheet_by_index(index_);
}

workbook::const_iterator &workbook::const_iterator::operator++()
{
    index_++;
    return *this;
}

workbook::const_iterator workbook::const_iterator::operator++(int)
{
    const_iterator old(wb_, index_);
    ++*this;
    return old;
}
    
bool workbook::const_iterator::operator==(const const_iterator &comparand) const
{
    return index_ == comparand.index_ && wb_ == comparand.wb_;
}
    
worksheet workbook::get_sheet_by_name(const std::string &name)
{
    auto title_equals = [&](detail::worksheet_impl &ws) { return worksheet(&ws).get_title() == name; };
    auto match = std::find_if(d_->worksheets_.begin(), d_->worksheets_.end(), title_equals);
    
    if(match != d_->worksheets_.end())
    {
        return worksheet(&*match);
    }

    return worksheet();
}

worksheet workbook::get_sheet_by_index(std::size_t index)
{
    return worksheet(&d_->worksheets_[index]);
}
    
const worksheet workbook::get_sheet_by_index(std::size_t index) const
{
    return worksheet(&d_->worksheets_.at(index));
}

worksheet workbook::get_active_sheet()
{
    return worksheet(&d_->worksheets_[d_->active_sheet_index_]);
}

bool workbook::has_named_range(const std::string &name) const
{
    for(auto worksheet : *this)
    {
        if(worksheet.has_named_range(name))
        {
            return true;
        }
    }
    return false;
}

worksheet workbook::create_sheet()
{
    if(d_->optimized_read_)
    {
        throw std::runtime_error("this is a read-only workbook");
    }
    
    std::string title = "Sheet1";
    int index = 1;

    while(get_sheet_by_name(title) != nullptr)
    {
        title = "Sheet" + std::to_string(++index);
    }

    d_->worksheets_.emplace_back(this, title);
    worksheet ws(&d_->worksheets_.back());
    ws.get_cell("A1");
    
    return ws;
}

void workbook::add_sheet(xlnt::worksheet worksheet)
{
    if(d_->optimized_read_)
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
    
    d_->worksheets_.emplace_back(*worksheet.d_);
}

void workbook::add_sheet(xlnt::worksheet worksheet, std::size_t index)
{
    add_sheet(worksheet);
    std::swap(d_->worksheets_[index], d_->worksheets_.back());
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

void workbook::create_named_range(const std::string &name, worksheet range_owner, const range_reference &reference)
{
    auto match = get_sheet_by_name(range_owner.get_title());
    if(match != nullptr)
    {
        match.create_named_range(name, reference);
        return;
    }
    throw std::runtime_error("worksheet isn't owned by this workbook");
}

void workbook::remove_named_range(const std::string &name)
{
    for(auto ws : *this)
    {
        if(ws.has_named_range(name))
        {
            ws.remove_named_range(name);
            return;
        }
    }
    
    throw std::runtime_error("named range not found");
}

range workbook::get_named_range(const std::string &name)
{
    for(auto ws : *this)
    {
        if(ws.has_named_range(name))
        {
            return ws.get_named_range(name);
        }
    }
    
    throw std::runtime_error("named range not found");
}

bool workbook::load(const std::vector<unsigned char> &data)
{
  std::ofstream tmp;
  tmp.open("/tmp/xlnt.xlsx", std::ios::out);
  for(auto c : data)
    {
      tmp.put(c);
    }
  load("/tmp/xlnt.xlsx");
  return true;
}

bool workbook::load(const std::string &filename)
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
    
    clear();
    
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

    return true;
}

void workbook::remove_sheet(worksheet ws)
{
    auto match_iter = std::find_if(d_->worksheets_.begin(), d_->worksheets_.end(), [=](detail::worksheet_impl &comp) { return worksheet(&comp) == ws; });

    if(match_iter == d_->worksheets_.end())
    {
        throw std::runtime_error("worksheet not owned by this workbook");
    }

    
    d_->worksheets_.erase(match_iter);
}

worksheet workbook::create_sheet(std::size_t index)
{
    create_sheet();
    
    if(index != d_->worksheets_.size())
    {
        std::swap(d_->worksheets_[index], d_->worksheets_.back());
    }
    
    return worksheet(&d_->worksheets_[index]);
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
        throw sheet_title_exception(title);
    }
    
    if(std::find_if(title.begin(), title.end(),
                    [](char c) { return c == '*' || c == ':' || c == '/' || c == '\\' || c == '?' || c == '[' || c == ']'; }) != title.end())
    {
        throw sheet_title_exception(title);
    }
    
    if(std::find_if(d_->worksheets_.begin(), d_->worksheets_.end(), [&](detail::worksheet_impl &ws) { return worksheet(&ws).get_title() == title; }) != d_->worksheets_.end())
    {
        throw std::runtime_error("sheet exists");
    }
    
    auto ws = create_sheet();
    ws.set_title(title);

    return ws;
}

workbook::iterator workbook::begin()
{
    return iterator(*this, 0);
}

workbook::iterator workbook::end()
{
    return iterator(*this, d_->worksheets_.size());
}

workbook::const_iterator workbook::cbegin() const
{
    return const_iterator(*this, 0);
}

workbook::const_iterator workbook::cend() const
{
    return const_iterator(*this, d_->worksheets_.size());
}

std::vector<std::string> workbook::get_sheet_names() const
{
    std::vector<std::string> names;
    
    for(auto ws : *this)
    {
        names.push_back(ws.get_title());
    }
    
    return names;
}

worksheet workbook::operator[](const std::string &name)
{
    return get_sheet_by_name(name);
}

worksheet workbook::operator[](int index)
{
    return worksheet(&d_->worksheets_[index]);
}

void workbook::clear()
{
    d_->worksheets_.clear();
}

bool workbook::save(std::vector<unsigned char> &data)
{
  save("/tmp/xlnt.xlsx");
  std::ifstream tmp;
  tmp.open("/tmp/xlnt.xlsx");
  auto char_data = std::vector<char>((std::istreambuf_iterator<char>(tmp)),
				     std::istreambuf_iterator<char>());
  data = std::vector<unsigned char>(char_data.begin(), char_data.end());
  return true;
}


bool workbook::save(const std::string &filename)
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
    
    std::vector<std::pair<std::string, std::pair<std::string, std::string>>> root_rels =
    {
        {"rId3", {"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml"}},
        {"rId2", {"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml"}},
        {"rId1", {"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml"}}
    };
    f.set_file_contents("_rels/.rels", writer::write_relationships(root_rels));
    
    std::vector<std::pair<std::string, std::pair<std::string, std::string>>> workbook_rels =
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

    return true;
}
    
bool workbook::operator==(const workbook &rhs) const
{
    return d_ == rhs.d_;
}
    
}
