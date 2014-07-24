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
#include "workbook/document_properties.hpp"
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

workbook_impl::workbook_impl() : active_sheet_index_(0), guess_types_(false), data_only_(false)
{
    
}

} // namespace detail
    
workbook::workbook() : d_(new detail::workbook_impl())
{
    create_sheet("Sheet");
    create_relationship("rId2", "sharedStrings.xml", relationship::type::shared_strings);
    create_relationship("rId3", "styles.xml", relationship::type::styles);
    create_relationship("rId4", "theme/theme1.xml", relationship::type::theme);
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
    create_relationship("rId" + std::to_string(d_->relationships_.size() + 1), "worksheets/sheet" + std::to_string(d_->worksheets_.size()) + ".xml", relationship::type::worksheet);
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
    auto content_types = reader::read_content_types(f);
    
    auto type = reader::determine_document_type(content_types);
    
    if(type != "excel")
    {
        throw std::runtime_error("unsupported document type: " + filename);
    }
    
    clear();
    
    auto workbook_relationships = reader::read_relationships(f, "xl/workbook.xml");

    for(auto relationship : workbook_relationships)
    {
        create_relationship(relationship.get_id(), relationship.get_target_uri(), relationship.get_type());
    }
    
    pugi::xml_document doc;
    doc.load(f.get_file_contents("xl/workbook.xml").c_str());
    
    auto root_node = doc.child("workbook");
    
    auto workbook_pr_node = root_node.child("workbookPr");
    get_properties().excel_base_date = (workbook_pr_node.attribute("date1904") != nullptr && workbook_pr_node.attribute("date1904").as_int() != 0) ? calendar::mac_1904 : calendar::windows_1900;
    
    auto sheets_node = root_node.child("sheets");
    
    std::vector<std::string> shared_strings;
    if(f.has_file("xl/sharedStrings.xml"))
    {
        shared_strings = xlnt::reader::read_shared_string(f.get_file_contents("xl/sharedStrings.xml"));
    }

    std::vector<int> number_format_ids;
    
    if(f.has_file("xl/styles.xml"))
    {
        pugi::xml_document styles_doc;
        styles_doc.load(f.get_file_contents("xl/styles.xml").c_str());
        auto stylesheet_node = styles_doc.child("styleSheet");
        auto cell_xfs_node = stylesheet_node.child("cellXfs");

        for(auto xf_node : cell_xfs_node.children("xf"))
        {
            number_format_ids.push_back(xf_node.attribute("numFmtId").as_int());
        }
    }
    
    for(auto sheet_node : sheets_node.children("sheet"))
    {
        std::string relation_id = sheet_node.attribute("r:id").as_string();
        auto ws = create_sheet(sheet_node.attribute("name").as_string());
        auto sheet_filename = get_relationship(relation_id).get_target_uri();
        xlnt::reader::read_worksheet(ws, f.get_file_contents(sheet_filename).c_str(), shared_strings, number_format_ids);
    }

    return true;
}

void workbook::set_guess_types(bool guess)
{
    d_->guess_types_ = guess;
}

bool workbook::get_guess_types() const
{
    return d_->guess_types_;
}

void workbook::create_relationship(const std::string &id, const std::string &target, relationship::type type)
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
    
    std::string unique_title = title;
    
    if(std::find_if(d_->worksheets_.begin(), d_->worksheets_.end(), [&](detail::worksheet_impl &ws) { return worksheet(&ws).get_title() == unique_title; }) != d_->worksheets_.end())
    {
        std::size_t suffix = 1;
        
        while(std::find_if(d_->worksheets_.begin(), d_->worksheets_.end(), [&](detail::worksheet_impl &ws) { return worksheet(&ws).get_title() == unique_title; }) != d_->worksheets_.end())
        {
            unique_title = title + std::to_string(suffix);
            suffix++;
        }
    }
    
    auto ws = create_sheet();
    ws.set_title(unique_title);

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
    d_->relationships_.clear();
    d_->active_sheet_index_ = 0;
    d_->drawings_.clear();
    d_->properties_ = document_properties();
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
    
    for(auto relationship : d_->relationships_)
    {
        if(relationship.get_type() == relationship::type::worksheet)
        {
            std::string sheet_index_string = relationship.get_target_uri().substr(16);
            std::size_t sheet_index = std::stoi(sheet_index_string.substr(0, sheet_index_string.find('.'))) - 1;
            f.set_file_contents("xl/" + relationship.get_target_uri(), writer::write_worksheet(get_sheet_by_index(sheet_index)));
        }
    }

    return true;
}

bool workbook::operator==(std::nullptr_t) const
{
    return d_.get() == nullptr;
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
	content_types.push_back({ false, "", "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" });
	return content_types;
}

document_properties &workbook::get_properties()
{
    return d_->properties_;
}

const document_properties &workbook::get_properties() const
{
    return d_->properties_;
}

void swap(workbook &left, workbook &right)
{
    using std::swap;
    swap(left.d_, right.d_);
    
    for(auto ws : left)
    {
        ws.set_parent(left);
    }
    
    for(auto ws : right)
    {
        ws.set_parent(right);
    }
}
    
workbook &workbook::operator=(workbook other)
{
    swap(*this, other);
    return *this;
}

workbook::workbook(workbook &&other) : workbook()
{
    swap(*this, other);
}
    
workbook::workbook(const workbook &other) : workbook()
{
    *d_.get() = *other.d_.get();
    
    for(auto ws : *this)
    {
        ws.set_parent(*this);
    }
}

bool workbook::get_data_only() const
{
    return d_->data_only_;
}

void workbook::set_data_only(bool data_only)
{
    d_->data_only_ = data_only;
}

}
