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
#include <detail/excel_serializer.hpp>
#include <detail/include_windows.hpp>
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
#include <xlnt/workbook/worksheet_iterator.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

workbook workbook::minimal()
{
	auto impl = new detail::workbook_impl();
	workbook wb(impl);

	wb.create_sheet();

	wb.d_->root_relationships_.push_back(relationship(relationship::type::office_document, "rId1", constants::part_workbook().to_string()));

	return wb;
}

workbook workbook::empty_excel()
{
	auto impl = new detail::workbook_impl();
	xlnt::workbook wb(impl);

	wb.d_->root_relationships_.push_back(relationship(relationship::type::core_properties, "rId1", constants::part_core().to_string()));
	wb.d_->root_relationships_.push_back(relationship(relationship::type::extended_properties, "rId2", constants::part_app().to_string()));
	wb.d_->root_relationships_.push_back(relationship(relationship::type::office_document, "rId3", constants::part_workbook().to_string()));

	wb.set_application("Microsoft Excel");
	wb.create_sheet();
	wb.add_format(format());
	wb.create_style("Normal");
	wb.set_theme(theme());

	auto &manifest = wb.d_->manifest_;
	manifest.add_override_type(constants::part_workbook(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
	manifest.add_override_type(constants::part_theme(), "application/vnd.openxmlformats-officedocument.theme+xml");
	manifest.add_override_type(constants::part_styles(), "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
	manifest.add_override_type(constants::part_core(), "application/vnd.openxmlformats-package.core-properties+xml");
	manifest.add_override_type(constants::part_app(), "application/vnd.openxmlformats-officedocument.extended-properties+xml");

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
	swap(*this, empty_excel());
}

workbook::workbook(detail::workbook_impl *impl) : d_(impl)
{
}

const worksheet workbook::get_sheet_by_name(const std::string &name) const
{
    for (auto &impl : d_->worksheets_)
    {
        if (impl.title_ == name)
        {
            return worksheet(&impl);
        }
    }

    throw key_not_found();
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
    std::string title = "Sheet";
    int index = 0;

    while (contains(title))
    {
        title = "Sheet" + std::to_string(++index);
    }

    auto sheet_id = d_->worksheets_.size() + 1;
    std::string sheet_filename = "sheet" + std::to_string(sheet_id) + ".xml";
	path sheet_path = constants::package_worksheets().append(sheet_filename);

    d_->worksheets_.push_back(detail::worksheet_impl(this, sheet_id, title));
    create_relationship("rId" + std::to_string(sheet_id),
                        "worksheets/" + sheet_filename,
                        relationship::type::worksheet);
    d_->manifest_.add_override_type(sheet_path,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");

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
    get_sheet_by_name(range_owner.get_title()).create_named_range(name, reference);
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
	return load(path(filename));
}

bool workbook::load(const path &filename)
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

void workbook::create_root_relationship(const std::string &id, const std::string &target, relationship::type type)
{
    d_->root_relationships_.push_back(relationship(type, id, target));
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

    d_->relationships_.erase(rel_iter);
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

worksheet workbook::operator[](const std::string &name)
{
    return get_sheet_by_name(name);
}

worksheet workbook::operator[](std::size_t index)
{
	return get_sheet_by_index(index);
}

void workbook::clear()
{
    d_->worksheets_.clear();
    d_->relationships_.clear();
    d_->active_sheet_index_ = 0;
    d_->manifest_.clear();
    clear_styles();
    clear_formats();
}

bool workbook::save(std::vector<unsigned char> &data)
{
    excel_serializer serializer(*this);
    serializer.save_virtual_workbook(data);

    return true;
}

bool workbook::save(const std::string &filename)
{
	return save(path(filename));
}

bool workbook::save(const path &filename)
{
	excel_serializer serializer(*this);
	serializer.save_workbook(filename);

	return true;
}

bool workbook::operator==(const workbook &rhs) const
{
    return d_.get() == rhs.d_.get();
}

bool workbook::operator!=(const workbook &rhs) const
{
    return d_.get() != rhs.d_.get();
}

const std::vector<relationship> &xlnt::workbook::get_relationships() const
{
    return d_->relationships_;
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

void workbook::set_code_name(const std::string & /*code_name*/)
{
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
	if (!d_->has_theme_)
	{
		bool has_theme_relationship = false;

		for (const auto &rel : get_relationships())
		{
			if (rel.get_type() == relationship::type::theme)
			{
				has_theme_relationship = true;
				break;
			}
		}

		if (!has_theme_relationship)
		{
			create_relationship(next_relationship_id(), "theme/theme1.xml", relationship::type::theme);
			d_->manifest_.add_override_type(constants::part_theme(), "application/vnd.openxmlformats-officedocument.spreadsheetml.theme+xml");
		}
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
	if (d_->stylesheet_.formats.empty())
	{
		bool has_style_relationship = false;

		for (const auto &rel : get_relationships())
		{
			if (rel.get_type() == relationship::type::styles)
			{
				has_style_relationship = true;
				break;
			}
		}

		if (!has_style_relationship)
		{
			create_relationship(next_relationship_id(), "styles.xml", relationship::type::styles);
			d_->manifest_.add_override_type(constants::part_styles(), "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
		}
	}

    return d_->stylesheet_.add_format(to_add);
}

std::size_t workbook::add_style(const style &to_add)
{
	if (d_->stylesheet_.formats.empty())
	{
		bool has_style_relationship = false;

		for (const auto &rel : get_relationships())
		{
			if (rel.get_type() == relationship::type::styles)
			{
				has_style_relationship = true;
				break;
			}
		}

		if (!has_style_relationship)
		{
			create_relationship(next_relationship_id(), "styles.xml", relationship::type::styles);
			d_->manifest_.add_override_type(constants::part_styles(), "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
		}
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

const std::vector<relationship> &workbook::get_root_relationships() const
{
    return d_->root_relationships_;
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
    if (d_->shared_strings_.empty())
    {
        bool has_shared_strings = false;

        for (const auto &rel : get_relationships())
        {
            if (rel.get_type() == relationship::type::shared_strings)
            {
                has_shared_strings = true;
                break;
            }
        }

        if (!has_shared_strings)
        {
            create_relationship(next_relationship_id(), "sharedStrings.xml", relationship::type::shared_strings);
            d_->manifest_.add_override_type(constants::part_shared_strings(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
        }
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

void workbook::set_thumbnail(const std::vector<std::uint8_t> &thumbnail)
{
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

std::string workbook::next_relationship_id() const
{
    std::size_t i = 1;
    
    while (std::find_if(d_->relationships_.begin(), d_->relationships_.end(),
        [i](const relationship &r) { return r.get_id() == "rId" + std::to_string(i); }) != d_->relationships_.end())
    {
        i++;
    }
    
    return "rId" + std::to_string(i);
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


} // namespace xlnt
