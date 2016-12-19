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

#ifdef _MSC_VER
#include <codecvt> // utf-8 -> utf-16 conversion
#endif

#include <detail/cell_impl.hpp>
#include <detail/constants.hpp>
#include <detail/excel_thumbnail.hpp>
#include <detail/vector_streambuf.hpp>
#include <detail/workbook_impl.hpp>
#include <detail/worksheet_impl.hpp>
#include <detail/xlsx_consumer.hpp>
#include <detail/xlsx_producer.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/packaging/relationship.hpp>
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
#include <xlnt/worksheet/header_footer.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace {

#ifdef _MSC_VER
std::wstring utf8_to_utf16(const std::string &utf8_string)
{
    std::wstring_convert<std::codecvt_utf8<wchar_t>> convert;
    return convert.from_bytes(utf8_string);
}

void open_stream(std::ifstream &stream, const std::wstring &path)
{
    stream.open(path, std::ios::binary);
}

void open_stream(std::ofstream &stream, const std::wstring &path)
{
    stream.open(path, std::ios::binary);
}

void open_stream(std::ifstream &stream, const std::string &path)
{
    open_stream(stream, utf8_to_utf16(path));
}

void open_stream(std::ofstream &stream, const std::string &path)
{
    open_stream(stream, utf8_to_utf16(path));
}
#else
void open_stream(std::ifstream &stream, const std::string &path)
{
    stream.open(path, std::ios::binary);
}

void open_stream(std::ofstream &stream, const std::string &path)
{
    stream.open(path, std::ios::binary);
}
#endif

} // namespace

namespace xlnt {

bool workbook::has_core_property(const std::string &property_name) const
{
    return d_->core_properties_.count(property_name) > 0;
}

template<>
void workbook::core_property(const std::string &property_name, const std::string value)
{
    d_->core_properties_[property_name] = value;
}

template<>
void workbook::core_property(const std::string &property_name, const char *value)
{
    d_->core_properties_[property_name] = value;
}

template<>
void workbook::core_property(const std::string &property_name, const datetime value)
{
    d_->core_properties_[property_name] = value.to_iso_string();
}

workbook workbook::minimal()
{
	auto impl = new detail::workbook_impl();
	workbook wb(impl);

	wb.d_->manifest_.register_default_type("rels", 
		"application/vnd.openxmlformats-package.relationships+xml");

	wb.d_->manifest_.register_override_type(path("/workbook.xml"),
		"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
	wb.d_->manifest_.register_relationship(uri("/"), relationship_type::office_document,
		uri("workbook.xml"), target_mode::internal);

	std::string title("1");
	wb.d_->worksheets_.push_back(detail::worksheet_impl(&wb, 1, title));

	wb.d_->manifest_.register_override_type(path("/sheet1.xml"),
		"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
	auto ws_rel = wb.d_->manifest_.register_relationship(uri("workbook.xml"),
		relationship_type::worksheet, uri("sheet1.xml"), target_mode::internal);
	wb.d_->sheet_title_rel_id_map_[title] = ws_rel;

	return wb;
}

workbook workbook::empty_excel()
{
	auto impl = new detail::workbook_impl();
	workbook wb(impl);

	wb.d_->manifest_.register_override_type(path("/xl/workbook.xml"),
		"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
	wb.d_->manifest_.register_relationship(uri("/"), relationship_type::office_document,
		uri("xl/workbook.xml"), target_mode::internal);

	wb.d_->manifest_.register_default_type("rels",
		"application/vnd.openxmlformats-package.relationships+xml");
	wb.d_->manifest_.register_default_type("xml", "application/xml");

	wb.thumbnail(excel_thumbnail(), "jpeg", "image/jpeg");

	wb.d_->manifest_.register_override_type(path("/docProps/core.xml"),
		"application/vnd.openxmlformats-package.core-properties+xml");
	wb.d_->manifest_.register_relationship(uri("/"), relationship_type::core_properties,
		uri("docProps/core.xml"), target_mode::internal);

	wb.d_->manifest_.register_override_type(path("/docProps/app.xml"),
		"application/vnd.openxmlformats-officedocument.extended-properties+xml");
	wb.d_->manifest_.register_relationship(uri("/"), relationship_type::extended_properties,
		uri("docProps/app.xml"), target_mode::internal);

	wb.core_property("Creator", "Microsoft Office User");
	wb.core_property("LastModifiedBy", "Microsoft Office User");
	wb.core_property("Created", datetime(2016, 8, 12, 3, 16, 56));
	wb.core_property("Modified", datetime(2016, 8, 12, 3, 17, 16));
	wb.core_property("Application", "Microsoft Macintosh Excel");
	wb.core_property("AppVersion", "15.0300");

    auto file_version = detail::workbook_impl::file_version_t{"xl", 6, 6, 26709 };
	wb.d_->file_version_ = file_version;

	auto ws = wb.create_sheet();

    page_margins margins;
    margins.left(0.7);
    margins.right(0.7);
    margins.top(0.75);
    margins.bottom(0.75);
    margins.header(0.3);
    margins.footer(0.3);
	ws.page_margins(margins);

    sheet_view view;
    ws.add_view(view);

    wb.d_->stylesheet_ = detail::stylesheet();
    auto &stylesheet = wb.d_->stylesheet_.get();

	auto default_border = border().side(border_side::bottom, border::border_property())
		.side(border_side::top, border::border_property())
		.side(border_side::start, border::border_property())
		.side(border_side::end, border::border_property())
		.side(border_side::diagonal, border::border_property());
	wb.d_->stylesheet_.get().borders.push_back(default_border);

    auto default_fill = fill(pattern_fill().type(pattern_fill_type::none));
	stylesheet.fills.push_back(default_fill);
    auto gray125_fill = pattern_fill().type(pattern_fill_type::gray125);
	stylesheet.fills.push_back(gray125_fill);

	auto default_font = font()
        .name("Calibri")
        .size(12)
        .scheme("minor")
        .family(2)
        .color(theme_color(1));
	stylesheet.fonts.push_back(default_font);

	wb.create_style("Normal").builtin_id(0)
        .border(default_border, false)
        .fill(default_fill, false)
        .font(default_font, false)
        .number_format(xlnt::number_format::general(), false);

    wb.create_format(true)
        .border(default_border, false)
        .fill(default_fill, false)
        .font(default_font, false)
        .number_format(xlnt::number_format::general(), false)
        .style("Normal");

	wb.theme(xlnt::theme());

	return wb;
}

workbook workbook::empty_libre_office()
{
	auto impl = new detail::workbook_impl();
	workbook wb(impl);

	wb.d_->manifest_.register_override_type(path("xl/workbook.xml"),
		"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
	wb.d_->manifest_.register_relationship(uri("/"), relationship_type::office_document,
		uri("xl/workbook.xml"), target_mode::internal);

	wb.d_->manifest_.register_override_type(path("docProps/core.xml"),
		"application/vnd.openxmlformats-package.core-properties+xml");
	wb.d_->manifest_.register_relationship(uri("/"), relationship_type::core_properties,
		uri("docProps/core.xml"), target_mode::internal);

	wb.d_->manifest_.register_override_type(path("docProps/app.xml"),
		"application/vnd.openxmlformats-officedocument.extended-properties+xml");
	wb.d_->manifest_.register_relationship(uri("/"), relationship_type::extended_properties,
		uri("docProps/app.xml"), target_mode::internal);

	wb.d_->manifest_.register_override_type(path("_rels/.rels"),
		"application/vnd.openxmlformats-package.relationships+xml");
	wb.d_->manifest_.register_override_type(path("xl/_rels/workbook.xml.rels"),
		"application/vnd.openxmlformats-package.relationships+xml");

	std::string title("Sheet1");
	wb.d_->worksheets_.push_back(detail::worksheet_impl(&wb, 1, title));

	wb.d_->manifest_.register_override_type(path("xl/worksheets/sheet1.xml"),
		"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
	auto ws_rel = wb.d_->manifest_.register_relationship(uri("xl/workbook.xml"),
		relationship_type::worksheet, uri("worksheets/sheet1.xml"), target_mode::internal);
	wb.d_->sheet_title_rel_id_map_[title] = ws_rel;

	auto ws = wb.sheet_by_index(0);
    
    sheet_view view;
    ws.add_view(view);

	page_margins margins;
	margins.left(0.7875);
	margins.right(0.7875);
	margins.top(1.05277777777778);
	margins.bottom(1.05277777777778);
	margins.header(0.7875);
	margins.footer(0.7875);
	ws.page_margins(margins);
	ws.page_setup(page_setup());
	ws.header_footer(xlnt::header_footer()
        .header(header_footer::location::center,
            formatted_text(text_run("&A").font("Times New Roman").size(12)))
        .footer(header_footer::location::center,
            formatted_text(text_run("Page &B").font("Times New Roman").size(12))));
	ws.add_column_properties(1, column_properties());

    wb.d_->stylesheet_ = detail::stylesheet();
    auto &stylesheet = wb.d_->stylesheet_.get();

	auto default_alignment = xlnt::alignment()
		.horizontal(horizontal_alignment::general)
		.vertical(vertical_alignment::bottom)
		.rotation(0)
		.wrap(false)
		.indent(0)
		.shrink(false);
	stylesheet.alignments.push_back(default_alignment);

	auto default_border = border().side(border_side::bottom, border::border_property())
		.side(border_side::top, border::border_property())
		.side(border_side::start, border::border_property())
		.side(border_side::end, border::border_property())
		.side(border_side::diagonal, border::border_property())
		.diagonal(diagonal_direction::neither);
	stylesheet.borders.push_back(default_border);

	auto default_fill = xlnt::fill(xlnt::pattern_fill().type(pattern_fill_type::none));
	stylesheet.fills.push_back(default_fill);

	auto gray125_fill = xlnt::fill(xlnt::pattern_fill().type(pattern_fill_type::gray125));
	stylesheet.fills.push_back(gray125_fill);

	auto default_font = font().name("Arial").size(10).family(2);
	stylesheet.fonts.push_back(default_font);

	auto second_font = font().name("Arial").size(10).family(0);
	stylesheet.fonts.push_back(second_font);

	auto default_number_format = xlnt::number_format();
	default_number_format.format_string("General");
	default_number_format.id(164);
	stylesheet.number_formats.push_back(default_number_format);

	auto default_protection = xlnt::protection()
		.locked(true)
		.hidden(false);
	stylesheet.protections.push_back(default_protection);

	wb.create_style("Normal").builtin_id(0).custom(false)
        .alignment(default_alignment, false)
        .border(default_border, true)
        .fill(default_fill, true)
        .font(default_font, true)
        .number_format(default_number_format, false)
        .protection(default_protection, false);

    wb.create_style("Comma").builtin_id(3).custom(false)
        .alignment(default_alignment, false)
        .border(default_border, true)
        .fill(default_fill, true)
        .font(second_font, true)
        .number_format(number_format::from_builtin_id(43), false)
        .protection(default_protection, false);

    wb.create_style("Comma [0]").builtin_id(6).custom(false)
        .alignment(default_alignment, false)
        .border(default_border, true)
        .fill(default_fill, true)
        .font(second_font, true)
        .number_format(number_format::from_builtin_id(41), false)
        .protection(default_protection, false);

    wb.create_style("Currency").builtin_id(4).custom(false)
        .alignment(default_alignment, false)
        .border(default_border, true)
        .fill(default_fill, true)
        .font(second_font, true)
        .number_format(number_format::from_builtin_id(44), false)
        .protection(default_protection, false);

    wb.create_style("Currency [0]").builtin_id(7).custom(false)
        .alignment(default_alignment, false)
        .border(default_border, true)
        .fill(default_fill, true)
        .font(second_font, true)
        .number_format(number_format::from_builtin_id(42), false)
        .protection(default_protection, false);

	wb.create_style("Percent").builtin_id(5).custom(false)
        .alignment(default_alignment, false)
        .border(default_border, true)
        .fill(default_fill, false)
        .font(second_font, true)
        .number_format(number_format::percentage(), false)
        .protection(default_protection, false);

	wb.create_format()
        .number_format(default_number_format, true)
        .alignment(default_alignment, true)
        .protection(default_protection, true)
        .style("Normal");

	wb.core_property("Application", "LibreOffice/5.1.4.2$Windows_x86 LibreOffice_project/f99d75f39f1c57ebdd7ffc5f42867c12031db97a");

	return wb;
}

workbook workbook::empty_numbers()
{
	return empty_excel();
}

workbook::workbook()
{
	auto wb_template = empty_excel();
	swap(*this, wb_template);
}

workbook::workbook(detail::workbook_impl *impl) : d_(impl)
{
    if (impl != nullptr)
    {
        if (d_->stylesheet_.is_set())
        {
            d_->stylesheet_->parent = this;
        }
    }
}

void workbook::register_app_properties_in_manifest()
{
	auto wb_rel = manifest().relationship(path("/"), relationship_type::office_document);

	if (!manifest().has_relationship(wb_rel.target().path(), relationship_type::extended_properties))
	{
		manifest().register_override_type(path("/docProps/app.xml"),
			"application/vnd.openxmlformats-officedocument.extended-properties+xml");
		manifest().register_relationship(uri("/"), relationship_type::extended_properties,
			uri("docProps/app.xml"), target_mode::internal);
	}
}

void workbook::register_core_properties_in_manifest()
{
	auto wb_rel = manifest().relationship(path("/"), relationship_type::office_document);

	if (!manifest().has_relationship(wb_rel.target().path(), relationship_type::core_properties))
	{
		manifest().register_override_type(path("/docProps/core.xml"),
			"application/vnd.openxmlformats-package.core-properties+xml");
		manifest().register_relationship(uri("/"), relationship_type::core_properties,
			uri("docProps/core.xml"), target_mode::internal);
	}
}

void workbook::register_shared_string_table_in_manifest()
{
	auto wb_rel = manifest().relationship(path("/"), relationship_type::office_document);

	if (!manifest().has_relationship(wb_rel.target().path(), relationship_type::shared_string_table))
	{
		manifest().register_override_type(constants::part_shared_strings(),
			"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
		manifest().register_relationship(wb_rel.target(),
			relationship_type::shared_string_table, uri("sharedStrings.xml"), target_mode::internal);
	}
}

void workbook::register_stylesheet_in_manifest()
{
	auto wb_rel = manifest().relationship(path("/"), relationship_type::office_document);

	if (!manifest().has_relationship(wb_rel.target().path(), relationship_type::stylesheet))
	{
		manifest().register_override_type(constants::part_styles(),
			"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
		manifest().register_relationship(wb_rel.target(),
			relationship_type::stylesheet, uri("styles.xml"), target_mode::internal);
	}
}

void workbook::register_theme_in_manifest()
{
	auto wb_rel = manifest().relationship(path("/"), relationship_type::office_document);

	if (!manifest().has_relationship(wb_rel.target().path(), relationship_type::theme))
	{
		manifest().register_override_type(constants::part_theme(),
			"application/vnd.openxmlformats-officedocument.theme+xml");
		manifest().register_relationship(wb_rel.target(),
			relationship_type::theme, uri("theme/theme1.xml"), target_mode::internal);
	}
}

void workbook::register_comments_in_manifest(worksheet ws)
{
	auto wb_rel = manifest().relationship(path("/"), relationship_type::office_document);
	auto ws_rel = manifest().relationship(wb_rel.target().path(), 
		d_->sheet_title_rel_id_map_.at(ws.title()));
	path ws_path(ws_rel.source().path().parent().append(ws_rel.target().path()));

	if (!manifest().has_relationship(ws_path, relationship_type::vml_drawing))
	{
		std::size_t file_number = 1;
		path filename("vmlDrawing1.vml");
        bool filename_exists = true;

		while (filename_exists)
		{
            filename_exists = false;

			for (auto current_ws_rel : manifest().relationships(wb_rel.target().path(), xlnt::relationship_type::worksheet))
			{
				path current_ws_path(current_ws_rel.source().path().parent().append(current_ws_rel.target().path()));
				if (!manifest().has_relationship(current_ws_path, xlnt::relationship_type::vml_drawing)) continue;

				for (auto current_ws_child_rel : manifest().relationships(current_ws_path, xlnt::relationship_type::vml_drawing))
				{
					if (current_ws_child_rel.target().path() == path("../drawings").append(filename))
					{
						filename_exists = true;
						break;
					}
				}
			}

            if (filename_exists)
            {
                file_number++;
                filename = path("vmlDrawing" + std::to_string(file_number) + ".vml");
            }
		}

		manifest().register_default_type("vml",
			"application/vnd.openxmlformats-officedocument.vmlDrawing");

		const path relative_path(path("../drawings").append(filename));
		manifest().register_relationship(uri(ws_path.string()),
			relationship_type::vml_drawing, uri(relative_path.string()), target_mode::internal);
	}

	if (!manifest().has_relationship(ws_path, relationship_type::comments))
	{
		std::size_t file_number = 1;
		path filename("comments1.xml");

		while (true)
		{
			if (!manifest().has_override_type(constants::package_xl().append(filename))) break;

			file_number++;
			filename = path("comments" + std::to_string(file_number) + ".xml");
		}

		const path absolute_path(constants::package_xl().append(filename));
		manifest().register_override_type(absolute_path,
			"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml");

		const path relative_path(path("..").append(filename));
		manifest().register_relationship(uri(ws_path.string()),
			relationship_type::comments, uri(relative_path.string()), target_mode::internal);
	}
}

const worksheet workbook::sheet_by_title(const std::string &title) const
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

worksheet workbook::sheet_by_title(const std::string &title)
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

worksheet workbook::sheet_by_index(std::size_t index)
{
    if (index >= d_->worksheets_.size())
    {
        throw invalid_parameter();
    }

	auto iter = d_->worksheets_.begin();

	for (std::size_t i = 0; i < index; ++i)
	{
        ++iter;
	}

	return worksheet(&*iter);
}

const worksheet workbook::sheet_by_index(std::size_t index) const
{
	auto iter = d_->worksheets_.begin();

	for (std::size_t i = 0; i < index; ++i, ++iter)
	{
	}

	return worksheet(&*iter);
}

worksheet workbook::sheet_by_id(std::size_t id)
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

const worksheet workbook::sheet_by_id(std::size_t id) const
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

worksheet workbook::active_sheet()
{
    return sheet_by_index(d_->active_sheet_index_.is_set()
        ? d_->active_sheet_index_.get() : 0);
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

	auto workbook_rel = d_->manifest_.relationship(path("/"), relationship_type::office_document);
    uri relative_sheet_uri(path("worksheets").append(sheet_filename).string());
    auto absolute_sheet_path = path("/xl").append(relative_sheet_uri.path());
	d_->manifest_.register_override_type(absolute_sheet_path,
		"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
	auto ws_rel = d_->manifest_.register_relationship(workbook_rel.target(),
		relationship_type::worksheet, relative_sheet_uri, target_mode::internal);
	d_->sheet_title_rel_id_map_[title] = ws_rel;

    return worksheet(&d_->worksheets_.back());
}

void workbook::copy_sheet(worksheet to_copy)
{
    if(to_copy.d_->parent_ != this) throw invalid_parameter();

    detail::worksheet_impl impl(*to_copy.d_);
    auto new_sheet = create_sheet();
    impl.title_ = new_sheet.title();
    *new_sheet.d_ = impl;
}

void workbook::copy_sheet(worksheet to_copy, std::size_t index)
{
    copy_sheet(to_copy);

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

std::size_t workbook::index(worksheet ws)
{
    auto match = std::find(begin(), end(), ws);

    if (match == end())
    {
		throw invalid_parameter();
    }

    return static_cast<std::size_t>(std::distance(begin(), match));
}

void workbook::create_named_range(const std::string &name, worksheet range_owner, const std::string &reference_string)
{
    create_named_range(name, range_owner, range_reference(reference_string));
}

void workbook::create_named_range(const std::string &name, worksheet range_owner, const range_reference &reference)
{
    sheet_by_title(range_owner.title()).create_named_range(name, reference);
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

	throw key_not_found();
}

range workbook::named_range(const std::string &name)
{
    for (auto ws : *this)
    {
        if (ws.has_named_range(name))
        {
            return ws.named_range(name);
        }
    }

    throw key_not_found();
}

void workbook::load(std::istream &stream)
{
	clear();
    detail::xlsx_consumer consumer(*this);
	consumer.read(stream);
}

void workbook::load(const std::vector<std::uint8_t> &data)
{
    if (data.size() < 22) // the shortest ZIP file is 22 bytes
    {
        throw xlnt::exception("file is empty or malformed");
    }

    xlnt::detail::vector_istreambuf data_buffer(data);
    std::istream data_stream(&data_buffer);
    load(data_stream);
}

void workbook::load(const std::string &filename)
{
	return load(path(filename));
}

void workbook::load(const path &filename)
{
	std::ifstream file_stream;
    open_stream(file_stream, filename.string());

    if (!file_stream.good())
    {
        throw xlnt::exception("file not found " + filename.string());
    }

    load(file_stream);
}

void workbook::load(const std::string &filename, const std::string &password)
{
	return load(path(filename), password);
}

void workbook::load(const path &filename, const std::string &password)
{
	std::ifstream file_stream;
    open_stream(file_stream, filename.string());

    if (!file_stream.good())
    {
        throw xlnt::exception("file not found " + filename.string());
    }

	return load(file_stream, password);
}

void workbook::load(const std::vector<std::uint8_t> &data, const std::string &password)
{
    if (data.size() < 22) // the shortest ZIP file is 22 bytes
    {
        throw xlnt::exception("file is empty or malformed");
    }

    xlnt::detail::vector_istreambuf data_buffer(data);
    std::istream data_stream(&data_buffer);
    load(data_stream, password);
}

void workbook::load(std::istream &stream, const std::string &password)
{
	clear();
	detail::xlsx_consumer consumer(*this);
	consumer.read(stream, password);
}

void workbook::save(std::vector<std::uint8_t> &data) const
{
    xlnt::detail::vector_ostreambuf data_buffer(data);
    std::ostream data_stream(&data_buffer);
    save(data_stream);
}

void workbook::save(const std::string &filename) const
{
	return save(path(filename));
}

void workbook::save(const path &filename) const
{
	std::ofstream file_stream;
    open_stream(file_stream, filename.string());
	save(file_stream);
}

void workbook::save(std::ostream &stream) const
{
	detail::xlsx_producer producer(*this);
	producer.write(stream);
}

void workbook::save(std::ostream &stream, const std::string &password)
{
    detail::xlsx_producer producer(*this);
    producer.write(stream, password);
}

#ifdef _MSC_VER
void workbook::save(const std::wstring &filename)
{
    std::ofstream file_stream;
    open_stream(file_stream, filename);
    save(file_stream);
}

void workbook::save(const std::wstring &filename, const std::string &password)
{
    std::ofstream file_stream;
    open_stream(file_stream, filename);
    save(file_stream, password);
}

void workbook::load(const std::wstring &filename)
{
    std::ifstream file_stream;
    open_stream(file_stream, filename);
    load(file_stream);
}

void workbook::load(const std::wstring &filename, const std::string &password)
{
    std::ifstream file_stream;
    open_stream(file_stream, filename);
    load(file_stream, password);
}
#endif

void workbook::remove_sheet(worksheet ws)
{
    auto match_iter = std::find_if(d_->worksheets_.begin(), d_->worksheets_.end(),
		[=](detail::worksheet_impl &comp) { return &comp == ws.d_; });

    if (match_iter == d_->worksheets_.end())
    {
		throw invalid_parameter();
    }

	auto ws_rel_id = d_->sheet_title_rel_id_map_.at(ws.title());
	auto wb_rel = d_->manifest_.relationship(xlnt::path("/"), xlnt::relationship_type::office_document);
	auto ws_rel = d_->manifest_.relationship(wb_rel.target().path(), ws_rel_id);
	d_->manifest_.unregister_override_type(ws_rel.target().path());
	d_->manifest_.unregister_relationship(wb_rel.target(), ws_rel_id);;
	d_->sheet_title_rel_id_map_.erase(ws.title());
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

	return sheet_by_index(index);
}

worksheet workbook::create_sheet_with_rel(const std::string &title, const relationship &rel)
{
    auto sheet_id = d_->worksheets_.size() + 1;
    d_->worksheets_.push_back(detail::worksheet_impl(this, sheet_id, title));

    auto workbook_rel = d_->manifest_.relationship(path("/"), relationship_type::office_document);
    auto sheet_absoulute_path = workbook_rel.target().path().parent().append(rel.target().path());
    d_->manifest_.register_override_type(sheet_absoulute_path,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
    auto ws_rel = d_->manifest_.register_relationship(workbook_rel.target(),
        relationship_type::worksheet, rel.target(), target_mode::internal);
    d_->sheet_title_rel_id_map_[title] = ws_rel;

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

std::vector<std::string> workbook::sheet_titles() const
{
    std::vector<std::string> names;

    for (auto ws : *this)
    {
        names.push_back(ws.title());
    }

    return names;
}

std::size_t workbook::sheet_count() const
{
	return d_->worksheets_.size();
}

worksheet workbook::operator[](const std::string &name)
{
    return sheet_by_title(name);
}

worksheet workbook::operator[](std::size_t index)
{
	return sheet_by_index(index);
}

void workbook::clear()
{
	*d_ = detail::workbook_impl();
    d_->stylesheet_.clear();
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
			ws.parent(left);
		}

		if (left.d_->stylesheet_.is_set())
		{
			left.d_->stylesheet_->parent = &left;
		}
	}

	if (right.d_ != nullptr)
	{
		for (auto ws : right)
		{
			ws.parent(right);
		}

		if (left.d_->stylesheet_.is_set())
		{
			right.d_->stylesheet_->parent = &right;
		}
	}
}

workbook &workbook::operator=(workbook other)
{
    swap(*this, other);
    d_->stylesheet_->parent = this;

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
        ws.parent(*this);
    }

    d_->stylesheet_->parent = this;
}

workbook::~workbook()
{
}

bool workbook::has_theme() const
{
	return d_->theme_.is_set();
}

const theme &workbook::theme() const
{
    return d_->theme_.get();
}

void workbook::theme(const class theme &value)
{
	register_theme_in_manifest();
	d_->theme_ = value;
}

std::vector<named_range> workbook::named_ranges() const
{
    std::vector<xlnt::named_range> named_ranges;

    for (auto ws : *this)
    {
        for (auto &ws_named_range : ws.d_->named_ranges_)
        {
            named_ranges.push_back(ws_named_range.second);
        }
    }

    return named_ranges;
}

format workbook::create_format(bool default_format)
{
	register_stylesheet_in_manifest();
    return d_->stylesheet_->create_format(default_format);
}

bool workbook::has_style(const std::string &name) const
{
	return d_->stylesheet_->has_style(name);
}

void workbook::clear_styles()
{
    apply_to_cells([](cell c) { c.clear_style(); });
}

void workbook::clear_formats()
{
    apply_to_cells([](cell c) { c.clear_format(); });
}

void workbook::apply_to_cells(std::function<void(cell)> f)
{
    for (auto ws : *this)
    {
        for (auto row = ws.lowest_row(); row <= ws.highest_row(); ++row)
        {
            for (auto column = ws.lowest_column(); column <= ws.highest_column(); ++column)
            {
                if (ws.has_cell(cell_reference(column, row)))
                {
                    f.operator()(ws.cell(cell_reference(column, row)));
                }
            }
        }
    }
}

format workbook::format(std::size_t format_index)
{
	return d_->stylesheet_->format(format_index);
}

const format workbook::format(std::size_t format_index) const
{
	return d_->stylesheet_->format(format_index);
}

manifest &workbook::manifest()
{
    return d_->manifest_;
}

const manifest &workbook::manifest() const
{
    return d_->manifest_;
}

std::vector<formatted_text> &workbook::shared_strings()
{
    return d_->shared_strings_;
}

const std::vector<formatted_text> &workbook::shared_strings() const
{
    return d_->shared_strings_;
}

void workbook::add_shared_string(const formatted_text &shared, bool allow_duplicates)
{
	register_shared_string_table_in_manifest();

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
        if(ws.title() == sheet_title) return true;
    }
    
    return false;
}

void workbook::thumbnail(const std::vector<std::uint8_t> &thumbnail,
	const std::string &extension, const std::string &content_type)
{
	if (!d_->manifest_.has_relationship(path("/"), relationship_type::thumbnail))
	{
        d_->manifest_.register_default_type(extension, content_type);
        d_->manifest_.register_relationship(uri("/"), relationship_type::thumbnail,
            uri("docProps/thumbnail.jpeg"), target_mode::internal);
	}

    auto thumbnail_rel = d_->manifest_.relationship(path("/"), relationship_type::thumbnail);
    d_->images_[thumbnail_rel.target().to_string()] = thumbnail;
}

const std::vector<std::uint8_t> &workbook::thumbnail() const
{
    auto thumbnail_rel = d_->manifest_.relationship(path("/"), relationship_type::thumbnail);
    return d_->images_.at(thumbnail_rel.target().to_string());
}

style workbook::create_style(const std::string &name)
{
	return d_->stylesheet_->create_style(name);
}

style workbook::style(const std::string &name)
{
	return d_->stylesheet_->style(name);
}

const style workbook::style(const std::string &name) const
{
	return d_->stylesheet_->style(name);
}

calendar workbook::base_date() const
{
	return d_->base_date_;
}

void workbook::base_date(calendar base_date)
{
	d_->base_date_ = base_date;
}

bool workbook::has_title() const
{
	return d_->title_.is_set();
}

std::string workbook::title() const
{
	return d_->title_.get();
}

void workbook::title(const std::string &title)
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
	return d_->view_.is_set();
}

workbook_view workbook::view() const
{
	if (!d_->view_.is_set())
	{
		throw invalid_attribute();
	}

	return d_->view_.get();
}

void workbook::view(const workbook_view &view)
{
	d_->view_ = view;
}

bool workbook::has_code_name() const
{
	return d_->code_name_.is_set();
}

std::string workbook::code_name() const
{
	if (has_code_name())
	{
		throw invalid_attribute();
	}

	return d_->code_name_.get();
}

void workbook::code_name(const std::string &code_name)
{
	d_->code_name_ = code_name;
}

bool workbook::has_file_version() const
{
	return d_->file_version_.is_set();
}

std::string workbook::app_name() const
{
	return d_->file_version_.get().app_name;
}

std::size_t workbook::last_edited() const
{
	return d_->file_version_.get().last_edited;
}

std::size_t workbook::lowest_edited() const
{
	return d_->file_version_.get().lowest_edited;
}

std::size_t workbook::rup_build() const
{
	return d_->file_version_.get().rup_build;
}

template<>
std::string workbook::core_property(const std::string &property_name) const
{
    return d_->core_properties_.at(property_name);
}

} // namespace xlnt
