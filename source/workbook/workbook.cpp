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
#include <functional>
#include <set>
#include <sstream>

#include <detail/cell_impl.hpp>
#include <detail/constants.hpp>
#include <detail/xlsx_consumer.hpp>
#include <detail/xlsx_producer.hpp>
#include <detail/workbook_impl.hpp>
#include <detail/worksheet_impl.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/packaging/zip_file.hpp>
#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/format.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/style.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/utils/path.hpp>
#include <xlnt/workbook/const_worksheet_iterator.hpp>
#include <xlnt/workbook/named_range.hpp>
#include <xlnt/workbook/theme.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/workbook/workbook_view.hpp>
#include <xlnt/workbook/worksheet_iterator.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

workbook workbook::minimal()
{
	auto impl = new detail::workbook_impl();
	workbook wb(impl);

	wb.d_->manifest_.register_default_type("rels", 
		"application/vnd.openxmlformats-package.relationships+xml");

	wb.d_->manifest_.register_override_type(path("workbook.xml"),
		"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
	wb.d_->manifest_.register_relationship(uri("/"), relationship::type::office_document,
		uri("workbook.xml"), target_mode::internal);

	std::string title("1");
	wb.d_->worksheets_.push_back(detail::worksheet_impl(&wb, 1, title));

	wb.d_->manifest_.register_override_type(path("sheet1.xml"),
		"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
	auto ws_rel = wb.d_->manifest_.register_relationship(uri("workbook.xml"),
		relationship::type::worksheet, uri("sheet1.xml"), target_mode::internal);
	wb.d_->sheet_title_rel_id_map_[title] = ws_rel;

	return wb;
}

workbook workbook::empty_excel()
{
	auto impl = new detail::workbook_impl();
	xlnt::workbook wb(impl);

	wb.d_->manifest_.register_override_type(path("xl/workbook.xml"),
		"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
	wb.d_->manifest_.register_relationship(uri("/"), relationship::type::office_document,
		uri("xl/workbook.xml"), target_mode::internal);

	wb.d_->manifest_.register_default_type("rels",
		"application/vnd.openxmlformats-package.relationships+xml");
	wb.d_->manifest_.register_default_type("xml", "application/xml");

	const std::vector<std::uint8_t> thumbnail = {
		0xff,0xd8,0xff,0xe0,0x00,0x10,0x4a,0x46,0x49,0x46,0x00,0x01,0x01,0x01,
		0x00,0x60,0x00,0x60,0x00,0x00,0xff,0xe1,0x00,0x16,0x45,0x78,0x69,0x66,
		0x00,0x00,0x49,0x49,0x2a,0x00,0x08,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
		0x00,0x00,0xff,0xdb,0x00,0x43,0x00,0x01,0x01,0x01,0x01,0x01,0x01,0x01,
		0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,
		0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,
		0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,
		0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,
		0x01,0xff,0xdb,0x00,0x43,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,
		0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,
		0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,
		0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,
		0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,0x01,
		0xff,0xc0,0x00,0x11,0x08,0x00,0x01,0x00,0x01,0x03,0x01,0x22,0x00,0x02,
		0x11,0x01,0x03,0x11,0x01,0xff,0xc4,0x00,0x15,0x00,0x01,0x01,0x00,0x00,
		0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x0a,
		0xff,0xc4,0x00,0x14,0x10,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
		0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xff,0xc4,0x00,0x14,0x01,0x01,
		0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
		0x00,0x00,0xff,0xc4,0x00,0x14,0x11,0x01,0x00,0x00,0x00,0x00,0x00,0x00,
		0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xff,0xda,0x00,0x0c,
		0x03,0x01,0x00,0x02,0x11,0x03,0x11,0x00,0x3f,0x00,0xbf,0x80,0x01,0xff,
		0xd9
	};

	wb.set_thumbnail(thumbnail, "jpeg", "image/jpeg");

	wb.d_->manifest_.register_override_type(path("docProps/core.xml"),
		"application/vnd.openxmlformats-package.coreproperties+xml");
	wb.d_->manifest_.register_relationship(uri("/"), relationship::type::core_properties,
		uri("docProps/core.xml"), target_mode::internal);

	wb.d_->manifest_.register_override_type(path("docProps/app.xml"),
		"application/vnd.openxmlformats-officedocument.extended-properties+xml");
	wb.d_->manifest_.register_relationship(uri("/"), relationship::type::extended_properties,
		uri("docProps/app.xml"), target_mode::internal);

	wb.set_creator("Microsoft Office User");
	wb.set_last_modified_by("Microsoft Office User");
	wb.set_created(datetime(2016, 8, 12, 3, 16, 56));
	wb.set_modified(datetime(2016, 8, 12, 3, 17, 16));
	wb.set_application("Microsoft Macintosh Excel");
	wb.set_app_version("15.0300");
	wb.set_absolute_path(path("/Users/thomas/Development/xlnt/tests/data/xlsx/"));
	wb.enable_x15();

	wb.d_->has_file_version_ = true;
	wb.d_->file_version_.app_name = "xl";
	wb.d_->file_version_.last_edited = 6;
	wb.d_->file_version_.lowest_edited = 6;
	wb.d_->file_version_.rup_build = 26709;

	wb.d_->has_properties_ = true;
	wb.d_->has_view_ = true;
	wb.d_->has_arch_id_ = true;
	wb.d_->has_calculation_properties_ = true;

	auto ws = wb.create_sheet();

	ws.enable_x14ac();
	ws.set_page_margins(page_margins());
	ws.d_->has_dimension_ = true;
	ws.d_->has_format_properties_ = true;
	ws.d_->has_view_ = true;

	wb.add_format(format());
	wb.create_style("Normal");
	wb.set_theme(theme());

	wb.d_->stylesheet_.format_styles.front() = "Normal";

	xlnt::fill gray125 = xlnt::fill::pattern(xlnt::pattern_fill::type::gray125);
	wb.d_->stylesheet_.fills.push_back(gray125);

	return wb;
}

workbook workbook::empty_libre_office()
{
	auto impl = new detail::workbook_impl();
	workbook wb(impl);

	return wb;
}

workbook workbook::empty_numbers()
{
	auto impl = new detail::workbook_impl();
	workbook wb(impl);

	return wb;
}

workbook::workbook()
{
	auto wb_template = empty_excel();
	swap(*this, wb_template);
}

workbook::workbook(detail::workbook_impl *impl) : d_(impl)
{
}

const worksheet workbook::get_sheet_by_title(const std::string &title) const
{
    for (auto &impl : d_->worksheets_)
    {
        if (impl.title_ == title)
        {
            return worksheet(&impl);
        }
    }

    throw key_not_found();
}

worksheet workbook::get_sheet_by_title(const std::string &title)
{
    for (auto &impl : d_->worksheets_)
    {
        if (impl.title_ == title)
        {
            return worksheet(&impl);
        }
    }

    throw key_not_found();
}

worksheet workbook::get_sheet_by_index(std::size_t index)
{
	auto iter = d_->worksheets_.begin();

	for (std::size_t i = 0; i < index; ++i, ++iter)
	{
	}

	return worksheet(&*iter);
}

const worksheet workbook::get_sheet_by_index(std::size_t index) const
{
	auto iter = d_->worksheets_.begin();

	for (std::size_t i = 0; i < index; ++i, ++iter)
	{
	}

	return worksheet(&*iter);
}

worksheet workbook::get_sheet_by_id(std::size_t id)
{
	for (auto &impl : d_->worksheets_)
	{
		if (impl.id_ == id)
		{
			return worksheet(&impl);
		}
	}

	throw key_not_found();
}

const worksheet workbook::get_sheet_by_id(std::size_t id) const
{
	for (auto &impl : d_->worksheets_)
	{
		if (impl.id_ == id)
		{
			return worksheet(&impl);
		}
	}

	throw key_not_found();
}

worksheet workbook::get_active_sheet()
{
    return get_sheet_by_index(d_->active_sheet_index_);
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

    while (contains(title))
    {
        title = "Sheet" + std::to_string(++index);
    }

    auto sheet_id = d_->worksheets_.size() + 1;
    std::string sheet_filename = "sheet" + std::to_string(sheet_id) + ".xml";

    d_->worksheets_.push_back(detail::worksheet_impl(this, sheet_id, title));

	auto workbook_rel = d_->manifest_.get_relationship(path("/"), relationship::type::office_document);
    uri sheet_uri(constants::package_worksheets().append(sheet_filename).string());
	d_->manifest_.register_override_type(sheet_uri.get_path(),
		"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
	auto ws_rel = d_->manifest_.register_relationship(workbook_rel.get_target(),
		relationship::type::worksheet, sheet_uri, target_mode::internal);
	d_->sheet_title_rel_id_map_[title] = ws_rel;

    return worksheet(&d_->worksheets_.back());
}

void workbook::copy_sheet(xlnt::worksheet worksheet)
{
    if(worksheet.d_->parent_ != this) throw xlnt::invalid_parameter();

    xlnt::detail::worksheet_impl impl(*worksheet.d_);
    auto new_sheet = create_sheet();
    impl.title_ = new_sheet.get_title();
    *new_sheet.d_ = impl;
}

void workbook::copy_sheet(xlnt::worksheet worksheet, std::size_t index)
{
    copy_sheet(worksheet);

    if (index != d_->worksheets_.size() - 1)
    {
		auto iter = d_->worksheets_.begin();

		for (std::size_t i = 0; i < index; ++i, ++iter)
		{
		}

		d_->worksheets_.insert(iter, d_->worksheets_.back());
		d_->worksheets_.pop_back();
    }
}

std::size_t workbook::get_index(xlnt::worksheet worksheet)
{
    auto match = std::find(begin(), end(), worksheet);

    if (match == end())
    {
        throw std::runtime_error("worksheet isn't owned by this workbook");
    }

    return std::distance(begin(), match);
}

void workbook::create_named_range(const std::string &name, worksheet range_owner, const std::string &reference_string)
{
    create_named_range(name, range_owner, range_reference(reference_string));
}

void workbook::create_named_range(const std::string &name, worksheet range_owner, const range_reference &reference)
{
    get_sheet_by_title(range_owner.get_title()).create_named_range(name, reference);
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

void workbook::load(std::istream &stream)
{
	clear();
    detail::xlsx_consumer consumer(*this);
	consumer.read(stream);
}

void workbook::load(const std::vector<unsigned char> &data)
{
	clear();
	detail::xlsx_consumer consumer(*this);
	consumer.read(data);
}

void workbook::load(const std::string &filename)
{
	return load(path(filename));
}

void workbook::load(const path &filename)
{
	clear();
	detail::xlsx_consumer consumer(*this);
	consumer.read(filename);
}

void workbook::save(std::vector<unsigned char> &data) const
{
	detail::xlsx_producer producer(*this);
	producer.write(data);
}

void workbook::save(const std::string &filename) const
{
	return save(path(filename));
}

void workbook::save(const path &filename) const
{
	detail::xlsx_producer producer(*this);
	producer.write(filename);
}

void workbook::save(std::ostream &stream) const
{
	detail::xlsx_producer producer(*this);
	producer.write(stream);
}

void workbook::set_guess_types(bool guess)
{
    d_->guess_types_ = guess;
}

bool workbook::get_guess_types() const
{
    return d_->guess_types_;
}

void workbook::remove_sheet(worksheet ws)
{
    auto match_iter = std::find_if(d_->worksheets_.begin(), d_->worksheets_.end(),
		[=](detail::worksheet_impl &comp) { return worksheet(&comp) == ws; });

    if (match_iter == d_->worksheets_.end())
    {
        throw std::runtime_error("worksheet not owned by this workbook");
    }

    auto sheet_filename = path("worksheets/sheet" + std::to_string(d_->worksheets_.size()) + ".xml");

    d_->worksheets_.erase(match_iter);
}

worksheet workbook::create_sheet(std::size_t index)
{
    create_sheet();

    if (index != d_->worksheets_.size() - 1)
    {
		auto iter = d_->worksheets_.begin();

		for (std::size_t i = 0; i < index; ++i, ++iter)
		{
		}

		d_->worksheets_.insert(iter, d_->worksheets_.back());
		d_->worksheets_.pop_back();
    }

	return get_sheet_by_index(index);
}

worksheet workbook::create_sheet_with_rel(const std::string &title, const relationship &rel)
{
    auto sheet_id = d_->worksheets_.size() + 1;
    d_->worksheets_.push_back(detail::worksheet_impl(this, sheet_id, title));

    return worksheet(&d_->worksheets_.back());
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

std::vector<std::string> workbook::get_sheet_titles() const
{
    std::vector<std::string> names;

    for (auto ws : *this)
    {
        names.push_back(ws.get_title());
    }

    return names;
}

std::size_t workbook::get_sheet_count() const
{
	return d_->worksheets_.size();
}

worksheet workbook::operator[](const std::string &name)
{
    return get_sheet_by_title(name);
}

worksheet workbook::operator[](std::size_t index)
{
	return get_sheet_by_index(index);
}

void workbook::clear()
{
	*d_ = detail::workbook_impl();
}

bool workbook::operator==(const workbook &rhs) const
{
    return d_.get() == rhs.d_.get();
}

bool workbook::operator!=(const workbook &rhs) const
{
    return d_.get() != rhs.d_.get();
}

void swap(workbook &left, workbook &right)
{
    using std::swap;
    swap(left.d_, right.d_);

	if (left.d_ != nullptr)
	{
		for (auto ws : left)
		{
			ws.set_parent(left);
		}
	}

	if (right.d_ != nullptr)
	{
		for (auto ws : right)
		{
			ws.set_parent(right);
		}
	}
}

workbook &workbook::operator=(workbook other)
{
    swap(*this, other);
    return *this;
}

workbook::workbook(workbook &&other) : workbook(nullptr)
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

workbook::~workbook()
{
}

bool workbook::get_data_only() const
{
    return d_->data_only_;
}

void workbook::set_data_only(bool data_only)
{
    d_->data_only_ = data_only;
}

bool workbook::has_theme() const
{
	return d_->has_theme_;
}

const theme &workbook::get_theme() const
{
    return d_->theme_;
}

void workbook::set_theme(const theme &value)
{
	auto workbook_rel = d_->manifest_.get_relationship(path("/"), relationship::type::office_document);

	if (!d_->manifest_.has_relationship(workbook_rel.get_target().get_path(), relationship::type::theme))
	{
		d_->manifest_.register_override_type(constants::part_theme(),
			"application/vnd.openxmlformats-officedocument.theme+xml");
		d_->manifest_.register_relationship(workbook_rel.get_target(),
			relationship::type::theme, uri(constants::part_theme().string()), target_mode::internal);
	}

	d_->has_theme_ = true;
	d_->theme_ = value;
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

std::size_t workbook::add_format(const format &to_add)
{
	auto workbook_rel = d_->manifest_.get_relationship(path("/"), relationship::type::office_document);

	if (!d_->manifest_.has_relationship(workbook_rel.get_target().get_path(), relationship::type::styles))
	{
		d_->manifest_.register_override_type(constants::part_styles(),
			"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
		d_->manifest_.register_relationship(workbook_rel.get_target(),
			relationship::type::styles, uri(constants::part_styles().string()), target_mode::internal);
	}

    return d_->stylesheet_.add_format(to_add);
}

std::size_t workbook::add_style(const style &to_add)
{
	auto workbook_rel = d_->manifest_.get_relationship(path("/"), relationship::type::office_document);

	if (!d_->manifest_.has_relationship(workbook_rel.get_target().get_path(), relationship::type::styles))
	{
		d_->manifest_.register_override_type(constants::part_styles(),
			"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
		d_->manifest_.register_relationship(workbook_rel.get_target(),
			relationship::type::styles, uri(constants::part_styles().string()), target_mode::internal);
	}

    return d_->stylesheet_.add_style(to_add);
}

bool workbook::has_style(const std::string &name) const
{
    return std::find_if(d_->stylesheet_.styles.begin(), d_->stylesheet_.styles.end(),
        [&](const style &s) { return s.get_name() == name; }) != d_->stylesheet_.styles.end();
}

std::size_t workbook::get_style_id(const std::string &name) const
{
    return std::distance(d_->stylesheet_.styles.begin(),
        std::find_if(d_->stylesheet_.styles.begin(), d_->stylesheet_.styles.end(),
            [&](const style &s) { return s.get_name() == name; }));
}

void workbook::clear_styles()
{
    d_->stylesheet_.styles.clear();
    apply_to_cells([](cell c) { c.clear_style(); });
}

void workbook::clear_formats()
{
    d_->stylesheet_.formats.clear();
    apply_to_cells([](cell c) { c.clear_format(); });
}

void workbook::apply_to_cells(std::function<void(cell)> f)
{
    for (auto ws : *this)
    {
        for (auto r : ws.iter_cells(true))
        {
            for (auto c : r)
            {
                f.operator()(c);
            }
        }
    }
}

format &workbook::get_format(std::size_t format_index)
{
    return d_->stylesheet_.formats.at(format_index);
}

const format &workbook::get_format(std::size_t format_index) const
{
    return d_->stylesheet_.formats.at(format_index);
}

manifest &workbook::get_manifest()
{
    return d_->manifest_;
}

const manifest &workbook::get_manifest() const
{
    return d_->manifest_;
}

std::vector<text> &workbook::get_shared_strings()
{
    return d_->shared_strings_;
}

const std::vector<text> &workbook::get_shared_strings() const
{
    return d_->shared_strings_;
}

void workbook::add_shared_string(const text &shared, bool allow_duplicates)
{
	auto workbook_rel = d_->manifest_.get_relationship(path("/"), relationship::type::office_document);

	if (!d_->manifest_.has_relationship(workbook_rel.get_target().get_path(), relationship::type::styles))
	{
		d_->manifest_.register_override_type(constants::part_shared_strings(),
			"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
		d_->manifest_.register_relationship(workbook_rel.get_target(),
			relationship::type::shared_string_table, uri(constants::part_shared_strings().string()), target_mode::internal);
	}

	if (!allow_duplicates)
	{
		//TODO: inefficient, use a set or something?
		for (auto &s : d_->shared_strings_)
		{
			if (s == shared) return;
		}
	}
    
    d_->shared_strings_.push_back(shared);
}

bool workbook::contains(const std::string &sheet_title) const
{
    for(auto ws : *this)
    {
        if(ws.get_title() == sheet_title) return true;
    }
    
    return false;
}

void workbook::set_thumbnail(const std::vector<std::uint8_t> &thumbnail,
	const std::string &extension, const std::string &content_type)
{
	if (!d_->manifest_.has_relationship(path("/"), relationship::type::thumbnail))
	{
        d_->manifest_.register_default_type(extension, content_type);
        d_->manifest_.register_relationship(uri("/"), relationship::type::thumbnail,
            uri("docProps/thumbnail.jpeg"), target_mode::internal);
	}

    d_->thumbnail_.assign(thumbnail.begin(), thumbnail.end());
}

const std::vector<std::uint8_t> &workbook::get_thumbnail() const
{
    return d_->thumbnail_;
}

style &workbook::create_style(const std::string &name)
{
    style style;
    style.set_name(name);
    
    d_->stylesheet_.styles.push_back(style);
    
    return d_->stylesheet_.styles.back();
}

style &workbook::get_style(const std::string &name)
{
    return *std::find_if(d_->stylesheet_.styles.begin(), d_->stylesheet_.styles.end(),
        [&name](const style &s) { return s.get_name() == name; });
}

const style &workbook::get_style(const std::string &name) const
{
    return *std::find_if(d_->stylesheet_.styles.begin(), d_->stylesheet_.styles.end(),
        [&name](const style &s) { return s.get_name() == name; });
}

style &workbook::get_style_by_id(std::size_t style_id)
{
    return d_->stylesheet_.styles.at(style_id);
}

const style &workbook::get_style_by_id(std::size_t style_id) const
{
    return d_->stylesheet_.styles.at(style_id);
}

std::string workbook::get_application() const
{
	return d_->application_;
}

void workbook::set_application(const std::string &application)
{
	d_->write_app_properties_ = true;
	d_->application_ = application;
}

calendar workbook::get_base_date() const
{
	return d_->base_date_;
}

void workbook::set_base_date(calendar base_date)
{
	d_->base_date_ = base_date;
}

std::string workbook::get_creator() const
{
	return d_->creator_;
}

void workbook::set_creator(const std::string &creator)
{
	d_->creator_ = creator;
}

std::string workbook::get_last_modified_by() const
{
	return d_->last_modified_by_;
}

void workbook::set_last_modified_by(const std::string &last_modified_by)
{
	d_->last_modified_by_ = last_modified_by;
}

datetime workbook::get_created() const
{
	return d_->created_;
}

void workbook::set_created(const datetime &when)
{
	d_->created_ = when;
}

datetime workbook::get_modified() const
{
	return d_->modified_;
}

void workbook::set_modified(const datetime &when)
{
	d_->modified_ = when;
}

int workbook::get_doc_security() const
{
	return d_->doc_security_;
}

void workbook::set_doc_security(int doc_security)
{
	d_->doc_security_ = doc_security;
}

bool workbook::get_scale_crop() const
{
	return d_->scale_crop_;
}

void workbook::set_scale_crop(bool scale_crop)
{
	d_->scale_crop_ = scale_crop;
}

std::string workbook::get_company() const
{
	return d_->company_;
}

void workbook::set_company(const std::string &company)
{
	d_->company_ = company;
}

bool workbook::links_up_to_date() const
{
	return d_->links_up_to_date_;
}

void workbook::set_links_up_to_date(bool links_up_to_date)
{
	d_->links_up_to_date_ = links_up_to_date;
}

bool workbook::is_shared_doc() const
{
	return d_->shared_doc_;
}

void workbook::set_shared_doc(bool shared_doc)
{
	d_->shared_doc_ = shared_doc;
}

bool workbook::hyperlinks_changed() const
{
	return d_->hyperlinks_changed_;
}

void workbook::set_hyperlinks_changed(bool hyperlinks_changed)
{
	d_->hyperlinks_changed_ = hyperlinks_changed;
}

std::string workbook::get_app_version() const
{
	return d_->app_version_;
}

void workbook::set_app_version(const std::string &version)
{
	d_->app_version_ = version;
}

std::string workbook::get_title() const
{
	return d_->title_;
}

void workbook::set_title(const std::string &title)
{
	d_->title_ = title;
}

detail::workbook_impl &workbook::impl()
{
	return *d_;
}

const detail::workbook_impl &workbook::impl() const
{
	return *d_;
}

bool workbook::has_view() const
{
	return d_->has_view_;
}

workbook_view workbook::get_view() const
{
	if (!d_->has_view_)
	{
		throw invalid_attribute();
	}

	return d_->view_;
}

void workbook::set_view(const workbook_view &view)
{
	d_->has_view_ = true;
	d_->view_ = view;
}

bool workbook::has_code_name() const
{
	return d_->has_code_name_;
}

std::string workbook::get_code_name() const
{
	if (!d_->has_code_name_)
	{
		throw invalid_attribute();
	}

	return d_->code_name_;
}

void workbook::set_code_name(const std::string &code_name)
{
	d_->code_name_ = code_name;
	d_->has_code_name_ = true;
}

bool workbook::x15_enabled() const
{
	return d_->x15_;
}

void workbook::enable_x15()
{
	d_->x15_ = true;
}

void workbook::disable_x15()
{
	d_->x15_ = false;
}

bool workbook::has_absolute_path() const
{
	return d_->has_absolute_path_;
}

path workbook::get_absolute_path() const
{
	return d_->absolute_path_;
}

void workbook::set_absolute_path(const path &absolute_path)
{
	d_->absolute_path_ = absolute_path;
	d_->has_absolute_path_ = true;
}

void workbook::clear_absolute_path()
{
	d_->absolute_path_ = path();
	d_->has_absolute_path_ = false;
}

bool workbook::has_properties() const
{
	return d_->has_properties_;
}

bool workbook::has_file_version() const
{
	return d_->has_file_version_;
}

std::string workbook::get_app_name() const
{
	return d_->file_version_.app_name;
}

std::size_t workbook::get_last_edited() const
{
	return d_->file_version_.last_edited;
}

std::size_t workbook::get_lowest_edited() const
{
	return d_->file_version_.lowest_edited;
}

std::size_t workbook::get_rup_build() const
{
	return d_->file_version_.rup_build;
}

bool workbook::has_calculation_properties() const
{
	return d_->has_calculation_properties_;
}

bool workbook::has_arch_id() const
{
	return d_->has_arch_id_;
}

} // namespace xlnt
