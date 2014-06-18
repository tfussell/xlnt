#include <algorithm>
#include <array>
#include <fstream>
#include <sstream>
#include <pugixml.hpp>

#ifdef _WIN32
#include <Windows.h>
#endif

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

static std::string CreateTemporaryFilename()
{
#ifdef _WIN32
    std::array<TCHAR, MAX_PATH> buffer;
    DWORD result = GetTempPath(static_cast<DWORD>(buffer.size()), buffer.data());
    if(result > MAX_PATH)
    {
        throw std::runtime_error("buffer is too small");
    }
    if(result == 0)
    {
        throw std::runtime_error("GetTempPath failed");
    }
    std::string directory(buffer.begin(), buffer.begin() + result);
    return directory + "xlnt.xlsx";
#else
    return "/tmp/xlsx.xlnt";
#endif
}

namespace xlnt {
namespace detail {
workbook_impl::workbook_impl() : active_sheet_index_(0), date_1904_(false)
{
    
}
} // namespace detail
    
workbook::workbook() : d_(new detail::workbook_impl())
{
    create_sheet("Sheet");
}

workbook::~workbook()
{
    clear();
}

workbook::iterator::iterator(workbook &wb, std::size_t index) : wb_(wb), index_(index)
{
    
}

workbook::iterator::iterator(const iterator &rhs) : wb_(rhs.wb_), index_(rhs.index_)
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

workbook::const_iterator::const_iterator(const const_iterator &rhs) : wb_(rhs.wb_), index_(rhs.index_)
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
    for(auto &impl : d_->worksheets_)
    {
        if(impl.title_ == name)
        {
            return worksheet(&impl);
        }
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
    std::string title = "Sheet1";
    int index = 1;

    while(get_sheet_by_name(title) != nullptr)
    {
        title = "Sheet" + std::to_string(++index);
    }

    d_->worksheets_.push_back(detail::worksheet_impl(this, title));
    return worksheet(&d_->worksheets_.back());
}

void workbook::add_sheet(xlnt::worksheet worksheet)
{
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

bool workbook::load(const std::istream &stream)
{
    std::string temp_file = CreateTemporaryFilename();
    
    std::ofstream tmp;
    tmp.open(temp_file, std::ios::out | std::ios::binary);
    tmp << stream.rdbuf();
    tmp.close();
    load(temp_file);
    std::remove(temp_file.c_str());
    return true;
}
    
bool workbook::load(const std::vector<unsigned char> &data)
{
    std::string temp_file = CreateTemporaryFilename();

    std::ofstream tmp;
    tmp.open(temp_file, std::ios::out | std::ios::binary);
    for(auto c : data)
    {
        tmp.put(c);
    }
    tmp.close();
    load(temp_file);
    std::remove(temp_file.c_str());
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

    for(auto relationship : workbook_relationships)
    {
        create_relationship(relationship.first, relationship.second.first, relationship.second.second);
    }
    
    pugi::xml_document doc;
    doc.load(f.get_file_contents("xl/workbook.xml").c_str());
    
    auto root_node = doc.child("workbook");
    
    auto workbook_pr_node = root_node.child("workbookPr");
    d_->date_1904_ = workbook_pr_node.attribute("date1904") != nullptr && workbook_pr_node.attribute("date1904").as_int() != 0;
    
    auto sheets_node = root_node.child("sheets");
    
    clear();
    
    std::vector<std::string> shared_strings;
    if(f.has_file("xl/sharedStrings.xml"))
    {
        shared_strings = xlnt::reader::read_shared_string(f.get_file_contents("xl/sharedStrings.xml"));
    }
    
    for(auto sheet_node : sheets_node.children("sheet"))
    {
        std::string relation_id = sheet_node.attribute("r:id").as_string();
        auto ws = create_sheet(sheet_node.attribute("name").as_string());
        std::string sheet_filename("xl/");
        sheet_filename += get_relationship(relation_id).get_target_uri();
        xlnt::reader::read_worksheet(ws, f.get_file_contents(sheet_filename).c_str(), shared_strings);
    }

    return true;
}

void workbook::create_relationship(const std::string &id, const std::string &target, const std::string &type)
{
    d_->relationships_.push_back(relationship(type, id, target));
}

relationship workbook::get_relationship(const std::string &id) const
{
    for(auto &rel : d_->relationships_)
    {
        if(rel.get_id() == id)
        {
            return rel;
        }
    }

    throw std::runtime_error("");
}
    
int workbook::get_base_year() const
{
    return d_->date_1904_ ? 1904 : 1900;
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

worksheet workbook::operator[](std::size_t index)
{
    return worksheet(&d_->worksheets_[index]);
}

void workbook::clear()
{
    d_->worksheets_.clear();
}

bool workbook::save(std::vector<unsigned char> &data)
{
    auto temp_file = CreateTemporaryFilename();
    save(temp_file);
    std::ifstream tmp;
    tmp.open(temp_file, std::ios::in | std::ios::binary);
    auto char_data = std::vector<char>((std::istreambuf_iterator<char>(tmp)),
        std::istreambuf_iterator<char>());
    data = std::vector<unsigned char>(char_data.begin(), char_data.end());
    tmp.close();
    std::remove(temp_file.c_str());
    return true;
}


bool workbook::save(const std::string &filename)
{
    zip_file f(filename, file_mode::create, file_access::write);

	f.set_file_contents("[Content_Types].xml", writer::write_content_types(*this));
    
    f.set_file_contents("_rels/.rels", writer::write_root_rels());
    f.set_file_contents("xl/_rels/workbook.xml.rels", writer::write_workbook_rels(*this));

    f.set_file_contents("xl/workbook.xml", writer::write_workbook(*this));

    return true;
}
    
bool workbook::operator==(const workbook &rhs) const
{
    return d_.get() == rhs.d_.get();
}

std::vector<relationship> xlnt::workbook::get_relationships() const
{
	return d_->relationships_;
}
 
std::vector<content_type> xlnt::workbook::get_content_types() const
{
	std::vector<content_type> content_types;
	content_types.push_back({ true, "xml", "", "application/xml" });
	content_types.push_back({ true, "rels", "", "application/vnd.openxmlformats-package.relationships+xml" });
	content_types.push_back({ false, "", "xl/workbook.xml", "applications/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" });
	return content_types;
}

}
