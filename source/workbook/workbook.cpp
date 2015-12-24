// Copyright (c) 2014-2016 Thomas Fussell
// Copyright (c) 2010-2015 openpyxl
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file
#include <algorithm>
#include <array>
#include <fstream>
#include <set>
#include <sstream>

#include <xlnt/drawing/drawing.hpp>
#include <xlnt/packaging/document_properties.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/packaging/zip_file.hpp>
#include <xlnt/serialization/encoding.hpp>
#include <xlnt/serialization/excel_serializer.hpp>
#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/styles/style.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/workbook/named_range.hpp>
#include <xlnt/workbook/theme.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/worksheet.hpp>

#include <detail/cell_impl.hpp>
#include <detail/include_windows.hpp>
#include <detail/workbook_impl.hpp>
#include <detail/worksheet_impl.hpp>

namespace xlnt {
namespace detail {

workbook_impl::workbook_impl()
    : active_sheet_index_(0), guess_types_(false), data_only_(false), next_custom_format_id_(164)
{
}

} // namespace detail

workbook::workbook() : d_(new detail::workbook_impl())
{
    create_sheet("Sheet");
    
    create_relationship("rId2", "sharedStrings.xml", relationship::type::shared_strings);
    create_relationship("rId3", "styles.xml", relationship::type::styles);
    create_relationship("rId4", "theme/theme1.xml", relationship::type::theme);
    
    add_number_format(number_format::general());

    d_->encoding_ = encoding::ascii;
    
    d_->manifest_.add_default_type("rels", "application/vnd.openxmlformats-package.relationships+xml");
    d_->manifest_.add_default_type("xml", "application/xml");
    
    d_->manifest_.add_override_type("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
    d_->manifest_.add_override_type("/xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml");
    d_->manifest_.add_override_type("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
    d_->manifest_.add_override_type("/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
    d_->manifest_.add_override_type("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
    d_->manifest_.add_override_type("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
}

workbook::workbook(encoding e) : workbook()
{
    d_->encoding_ = e;
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
    for (auto &impl : d_->worksheets_)
    {
        if (impl.title_ == name)
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
    for (auto worksheet : *this)
    {
        if (worksheet.has_named_range(name))
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

    while (get_sheet_by_name(title) != nullptr)
    {
        title = "Sheet" + std::to_string(++index);
    }
    
    std::string sheet_filename = "worksheets/sheet" + std::to_string(d_->worksheets_.size() + 1) + ".xml";

    d_->worksheets_.push_back(detail::worksheet_impl(this, title));
    create_relationship("rId" + std::to_string(d_->relationships_.size() + 1),
                        sheet_filename,
                        relationship::type::worksheet);
    
    d_->manifest_.add_override_type("/xl/" + sheet_filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");

    return worksheet(&d_->worksheets_.back());
}

void workbook::add_sheet(xlnt::worksheet worksheet)
{
    for (auto ws : *this)
    {
        if (worksheet == ws)
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
    for (auto ws : *this)
    {
        if (worksheet == ws)
        {
            return i;
        }
        i++;
    }
    throw std::runtime_error("worksheet isn't owned by this workbook");
}

void workbook::create_named_range(const std::string &name, worksheet range_owner, const std::string &reference_string)
{
    create_named_range(name, range_owner, range_reference(reference_string));
}

void workbook::create_named_range(const std::string &name, worksheet range_owner, const range_reference &reference)
{
    auto match = get_sheet_by_name(range_owner.get_title());
    if (match != nullptr)
    {
        match.create_named_range(name, reference);
        return;
    }
    throw std::runtime_error("worksheet isn't owned by this workbook");
}

void workbook::remove_named_range(const std::string &name)
{
    for (auto ws : *this)
    {
        if (ws.has_named_range(name))
        {
            ws.remove_named_range(name);
            return;
        }
    }

    throw std::runtime_error("named range not found");
}

range workbook::get_named_range(const std::string &name)
{
    for (auto ws : *this)
    {
        if (ws.has_named_range(name))
        {
            return ws.get_named_range(name);
        }
    }

    throw std::runtime_error("named range not found");
}

bool workbook::load(std::istream &stream)
{
    excel_serializer serializer_(*this);
    serializer_.load_stream_workbook(stream);

    return true;
}

bool workbook::load(const std::vector<unsigned char> &data)
{
    excel_serializer serializer_(*this);
    serializer_.load_virtual_workbook(data);

    return true;
}

bool workbook::load(const std::string &filename)
{
    excel_serializer serializer_(*this);
    serializer_.load_workbook(filename);

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
    for (auto &rel : d_->relationships_)
    {
        if (rel.get_id() == id)
        {
            return rel;
        }
    }

    throw std::runtime_error("");
}

void workbook::remove_sheet(worksheet ws)
{
    auto match_iter = std::find_if(d_->worksheets_.begin(), d_->worksheets_.end(),
                                   [=](detail::worksheet_impl &comp) { return worksheet(&comp) == ws; });

    if (match_iter == d_->worksheets_.end())
    {
        throw std::runtime_error("worksheet not owned by this workbook");
    }

    auto sheet_filename = "worksheets/sheet" + std::to_string(d_->worksheets_.size()) + ".xml";
    auto rel_iter = std::find_if(d_->relationships_.begin(), d_->relationships_.end(),
                                 [=](relationship &r) { return r.get_target_uri() == sheet_filename; });

    if (rel_iter == d_->relationships_.end())
    {
        throw std::runtime_error("no matching rel found");
    }

    d_->relationships_.erase(rel_iter);
    d_->worksheets_.erase(match_iter);
}

worksheet workbook::create_sheet(std::size_t index)
{
    create_sheet();

    if (index != d_->worksheets_.size() - 1)
    {
        std::swap(d_->worksheets_.back(), d_->worksheets_[index]);
        d_->worksheets_.pop_back();
    }

    return worksheet(&d_->worksheets_[index]);
}

// TODO: There should be a better way to do this...
std::size_t workbook::index_from_ws_filename(const std::string &ws_filename)
{
    std::string sheet_index_string(ws_filename);
    sheet_index_string = sheet_index_string.substr(0, sheet_index_string.find('.'));
    sheet_index_string = sheet_index_string.substr(sheet_index_string.find_last_of('/'));
    auto iter = sheet_index_string.end();
    iter--;
    while (isdigit(*iter))
        iter--;
    auto first_digit = static_cast<std::size_t>(iter - sheet_index_string.begin());
    sheet_index_string = sheet_index_string.substr(first_digit + 1);
    auto sheet_index = static_cast<std::size_t>(std::stoll(sheet_index_string) - 1);
    return sheet_index;
}

worksheet workbook::create_sheet(const std::string &title, const relationship &rel)
{
    d_->worksheets_.push_back(detail::worksheet_impl(this, title));

    auto index = index_from_ws_filename(rel.get_target_uri());
    if (index != d_->worksheets_.size() - 1)
    {
        std::swap(d_->worksheets_.back(), d_->worksheets_[index]);
        d_->worksheets_.pop_back();
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
    if (title.length() > 31)
    {
        throw sheet_title_exception(title);
    }

    if (std::find_if(title.begin(), title.end(), [](char c) {
            return c == '*' || c == ':' || c == '/' || c == '\\' || c == '?' || c == '[' || c == ']';
        }) != title.end())
    {
        throw sheet_title_exception(title);
    }

    std::string unique_title = title;

    if (std::find_if(d_->worksheets_.begin(), d_->worksheets_.end(), [&](detail::worksheet_impl &ws) {
            return worksheet(&ws).get_title() == unique_title;
        }) != d_->worksheets_.end())
    {
        std::size_t suffix = 1;

        while (std::find_if(d_->worksheets_.begin(), d_->worksheets_.end(), [&](detail::worksheet_impl &ws) {
                   return worksheet(&ws).get_title() == unique_title;
               }) != d_->worksheets_.end())
        {
            unique_title = title + std::to_string(suffix);
            suffix++;
        }
    }

    auto ws = create_sheet();
    ws.set_title(unique_title);

    return ws;
}

encoding workbook::get_encoding() const
{
    return d_->encoding_;
}

workbook::iterator workbook::begin()
{
    return iterator(*this, 0);
}

workbook::iterator workbook::end()
{
    return iterator(*this, d_->worksheets_.size());
}

workbook::const_iterator workbook::begin() const
{
    return cbegin();
}

workbook::const_iterator workbook::end() const
{
    return cend();
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

    for (auto ws : *this)
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
    excel_serializer serializer(*this);
    serializer.save_virtual_workbook(data);

    return true;
}

bool workbook::save(const std::string &filename)
{
    excel_serializer serializer(*this);
    serializer.save_workbook(filename);

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

const std::vector<relationship> &xlnt::workbook::get_relationships() const
{
    return d_->relationships_;
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

    for (auto ws : left)
    {
        ws.set_parent(left);
    }

    for (auto ws : right)
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

    for (auto ws : *this)
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

void workbook::add_border(const xlnt::border &border_)
{
    d_->borders_.push_back(border_);
}

void workbook::add_fill(const fill &fill_)
{
    d_->fills_.push_back(fill_);
}

void workbook::add_font(const font &font_)
{
    d_->fonts_.push_back(font_);
}

void workbook::add_number_format(const number_format &number_format_)
{
    if (d_->number_formats_.size() == 1 && d_->number_formats_.front() == number_format::general())
    {
        d_->number_formats_.front() = number_format_;
    }
    else
    {
        d_->number_formats_.push_back(number_format_);
    }
    
    if(!number_format_.has_id())
    {
        d_->number_formats_.back().set_id(d_->next_custom_format_id_++);
    }
}

void workbook::set_code_name(const std::string & /*code_name*/)
{
}

bool workbook::has_loaded_theme() const
{
    return false;
}

const theme &workbook::get_loaded_theme() const
{
    return d_->theme_;
}

std::vector<named_range> workbook::get_named_ranges() const
{
    std::vector<named_range> named_ranges;

    for (auto ws : *this)
    {
        for (auto &ws_named_range : ws.d_->named_ranges_)
        {
            named_ranges.push_back(ws_named_range.second);
        }
    }

    return named_ranges;
}

std::size_t workbook::add_style(const xlnt::style &style_)
{
    d_->styles_.push_back(style_);
    d_->styles_.back().id_ = d_->styles_.size() - 1;

    return d_->styles_.back().id_;
}

color workbook::add_indexed_color(const color &rgb_color)
{
	std::size_t index = 0;

	for (auto &c : d_->colors_)
	{
		if (c.get_rgb_string() == rgb_color.get_rgb_string())
		{
			return color(color::type::indexed, index);
		}

		index++;
	}

	d_->colors_.push_back(rgb_color);

	return color(color::type::indexed, index);
}

color workbook::get_indexed_color(const color &indexed_color) const
{
	return d_->colors_.at(indexed_color.get_index());
}

const number_format &workbook::get_number_format(std::size_t style_id) const
{
    auto number_format_id = d_->styles_[style_id].number_format_id_;

    for (const auto &number_format_ : d_->number_formats_)
    {
        if (number_format_.has_id() && number_format_.get_id() == number_format_id)
        {
            return number_format_;
        }
    }

    auto nf = number_format::from_builtin_id(number_format_id);
    d_->number_formats_.push_back(nf);

    return d_->number_formats_.back();
}

const font &workbook::get_font(std::size_t font_id) const
{
    return d_->fonts_[font_id];
}

std::size_t workbook::set_font(const font &font_, std::size_t style_id)
{
    auto match = std::find(d_->fonts_.begin(), d_->fonts_.end(), font_);
    std::size_t font_id = 0;

    if (match == d_->fonts_.end())
    {
        d_->fonts_.push_back(font_);
        font_id = d_->fonts_.size() - 1;
    }
    else
    {
        font_id = match - d_->fonts_.begin();
    }

    if (d_->styles_.empty())
    {
        style new_style;

        new_style.id_ = 0;
        new_style.border_id_ = 0;
        new_style.fill_id_ = 0;
        new_style.font_id_ = font_id;
        new_style.font_apply_ = true;
        new_style.number_format_id_ = 0;

        if (d_->borders_.empty())
        {
            d_->borders_.push_back(new_style.get_border());
        }

        if (d_->fills_.empty())
        {
            d_->fills_.push_back(new_style.get_fill());
        }

        if (d_->number_formats_.empty())
        {
            d_->number_formats_.push_back(new_style.get_number_format());
        }

        d_->styles_.push_back(new_style);

        return 0;
    }

    // If the style is unchanged, just return it.
    auto &existing_style = d_->styles_[style_id];
    existing_style.font_apply_ = true;

    if (font_id == existing_style.font_id_)
    {
        // no change
        return style_id;
    }

    // Make a new style with this format.
    auto new_style = existing_style;

    new_style.font_id_ = font_id;
    new_style.font_ = font_;

    // Check if the new style is already applied to a different cell. If so, reuse it.
    auto style_match = std::find(d_->styles_.begin(), d_->styles_.end(), new_style);

    if (style_match != d_->styles_.end())
    {
        return style_match->get_id();
    }

    // No match found, so add it.
    new_style.id_ = d_->styles_.size();
    d_->styles_.push_back(new_style);

    return new_style.id_;
}

const fill &workbook::get_fill(std::size_t fill_id) const
{
    return d_->fills_[fill_id];
}

std::size_t workbook::set_fill(const fill & /*fill_*/, std::size_t style_id)
{
    return style_id;
}

const border &workbook::get_border(std::size_t border_id) const
{
    return d_->borders_[border_id];
}

std::size_t workbook::set_border(const border & /*border_*/, std::size_t style_id)
{
    return style_id;
}

const alignment &workbook::get_alignment(std::size_t style_id) const
{
    return d_->styles_[style_id].alignment_;
}

std::size_t workbook::set_alignment(const alignment & /*alignment_*/, std::size_t style_id)
{
    return style_id;
}

const protection &workbook::get_protection(std::size_t style_id) const
{
    return d_->styles_[style_id].protection_;
}

std::size_t workbook::set_protection(const protection & /*protection_*/, std::size_t style_id)
{
    return style_id;
}

bool workbook::get_pivot_button(std::size_t style_id) const
{
    return d_->styles_[style_id].pivot_button_;
}

bool workbook::get_quote_prefix(std::size_t style_id) const
{
    return d_->styles_[style_id].quote_prefix_;
}

//TODO: this is terrible!
std::size_t workbook::set_number_format(const xlnt::number_format &format, std::size_t style_id)
{
    auto match = std::find(d_->number_formats_.begin(), d_->number_formats_.end(), format);
    std::size_t format_id = 0;

    if (match == d_->number_formats_.end())
    {
        d_->number_formats_.push_back(format);

        if (!format.has_id())
        {
            d_->number_formats_.back().set_id(d_->next_custom_format_id_++);
        }

        format_id = d_->number_formats_.back().get_id();
    }
    else
    {
        format_id = match->get_id();
    }

    if (d_->styles_.empty())
    {
        style new_style;

        new_style.id_ = 0;
        new_style.border_id_ = 0;
        new_style.fill_id_ = 0;
        new_style.font_id_ = 0;
        new_style.number_format_id_ = format_id;
        new_style.number_format_apply_ = true;

        if (d_->borders_.empty())
        {
            d_->borders_.push_back(new_style.get_border());
        }

        if (d_->fills_.empty())
        {
            d_->fills_.push_back(new_style.get_fill());
        }

        if (d_->fonts_.empty())
        {
            d_->fonts_.push_back(new_style.get_font());
        }

        d_->styles_.push_back(new_style);

        return 0;
    }

    // If the style is unchanged, just return it.
    auto existing_style = d_->styles_[style_id];
    existing_style.number_format_apply_ = true;

    if (format_id == existing_style.number_format_id_)
    {
        // no change
        return style_id;
    }

    // Make a new style with this format.
    auto new_style = existing_style;

    new_style.number_format_id_ = format_id;
    new_style.number_format_ = format;

    // Check if the new style is already applied to a different cell. If so, reuse it.
    auto style_match = std::find(d_->styles_.begin(), d_->styles_.end(), new_style);

    if (style_match != d_->styles_.end())
    {
        return style_match->get_id();
    }

    // No match found, so add it.
    new_style.id_ = d_->styles_.size();
    d_->styles_.push_back(new_style);

    return new_style.id_;
}

const std::vector<style> &workbook::get_styles() const
{
    return d_->styles_;
}

const std::vector<number_format> &workbook::get_number_formats() const
{
    return d_->number_formats_;
}

const std::vector<font> &workbook::get_fonts() const
{
    return d_->fonts_;
}

const std::vector<border> &workbook::get_borders() const
{
    return d_->borders_;
}

const std::vector<fill> &workbook::get_fills() const
{
    return d_->fills_;
}

void workbook::add_color(const color &rgb)
{
    d_->colors_.push_back(rgb);
}

const std::vector<color> &workbook::get_colors() const
{
    return d_->colors_;
}

manifest &workbook::get_manifest()
{
    return d_->manifest_;
}

const manifest &workbook::get_manifest() const
{
    return d_->manifest_;
}

const std::vector<relationship> &workbook::get_root_relationships() const
{
    if (d_->root_relationships_.empty())
    {
        d_->root_relationships_.push_back(
            relationship(relationship::type::core_properties, "rId1", "docProps/core.xml"));
        d_->root_relationships_.push_back(
            relationship(relationship::type::extended_properties, "rId2", "docProps/app.xml"));
        d_->root_relationships_.push_back(relationship(relationship::type::office_document, "rId3", "xl/workbook.xml"));
    }

    return d_->root_relationships_;
}

std::vector<std::string> &workbook::get_shared_strings()
{
    return d_->shared_strings_;
}

const std::vector<std::string> &workbook::get_shared_strings() const
{
    return d_->shared_strings_;
}

void workbook::add_shared_string(const std::string &shared)
{
    //TODO: inefficient, use a set or something?
    for(auto &s : d_->shared_strings_)
    {
        if(s == shared) return;
    }
    
    d_->shared_strings_.push_back(shared);
}

} // namespace xlnt
