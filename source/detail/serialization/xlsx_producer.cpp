// Copyright (c) 2014-2017 Thomas Fussell
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

#include <cmath>
#include <numeric> // for std::accumulate
#include <string>
#include <unordered_set>

#include <detail/constants.hpp>
#include <detail/implementations/workbook_impl.hpp>
#include <detail/header_footer/header_footer_code.hpp>
#include <detail/serialization/custom_value_traits.hpp>
#include <detail/serialization/vector_streambuf.hpp>
#include <detail/serialization/xlsx_producer.hpp>
#include <detail/serialization/zstream.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/utils/path.hpp>
#include <xlnt/utils/scoped_enum_hash.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/workbook/workbook_view.hpp>
#include <xlnt/worksheet/header_footer.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace {

/// <summary>
/// Returns true if d is exactly equal to an integer.
/// </summary>
bool is_integral(double d)
{
    return std::fabs(d - static_cast<double>(static_cast<long long int>(d))) == 0.0;
}

std::vector<std::pair<std::string, std::string>> core_property_namespace(xlnt::core_property type)
{
    using xlnt::core_property;
    using xlnt::constants;

    if (type == core_property::created
        || type == core_property::modified)
    {
        return {{constants::ns("dcterms"), "dcterms"},
            {constants::ns("xsi"), "xsi"}};
    }
    else if (type == core_property::title
        || type == core_property::subject
        || type == core_property::creator
        || type == core_property::description)
    {
        return {{constants::ns("dc"), "dc"}};
    }
    else if (type == core_property::keywords)
    {
        return {{constants::ns("core-properties"), "cp"},
            {constants::ns("vt"), "vt"}};
    }

    return {{constants::ns("core-properties"), "cp"}};
}

} // namespace

namespace xlnt {
namespace detail {

xlsx_producer::xlsx_producer(const workbook &target)
    : source_(target),
      current_part_stream_(nullptr)
{
}

xlsx_producer::~xlsx_producer()
{
    end_part();
    archive_.reset();
}

void xlsx_producer::write(std::ostream &destination)
{
    archive_.reset(new ozstream(destination));
    populate_archive(false);
}

void xlsx_producer::open(std::ostream &destination)
{
    archive_.reset(new ozstream(destination));
    populate_archive(true);
}

cell xlsx_producer::add_cell(const cell_reference &ref)
{
    current_cell_->column_ = ref.column();
    current_cell_->row_ = ref.row();

    return cell(current_cell_);
}

worksheet xlsx_producer::add_worksheet(const std::string &title)
{
    current_worksheet_->title_ = title;
    return worksheet(current_worksheet_);
}

// Part Writing Methods

void xlsx_producer::populate_archive(bool streaming)
{
    streaming_ = streaming;

    write_content_types();

    const auto root_rels = source_.manifest().relationships(path("/"));
    write_relationships(root_rels, path("/"));

    for (auto &rel : root_rels)
    {
        // thumbnail is binary content so we don't want to open an xml serializer stream
        if (rel.type() == relationship_type::thumbnail)
        {
            write_image(rel.target().path());
            continue;
        }

        begin_part(rel.target().path());

        if (rel.type() == relationship_type::core_properties)
        {
            write_core_properties(rel);
        }
        else if (rel.type() == relationship_type::extended_properties)
        {
            write_extended_properties(rel);
        }
        else if (rel.type() == relationship_type::custom_properties)
        {
            write_custom_properties(rel);
        }
        else if (rel.type() == relationship_type::office_document)
        {
            write_workbook(rel);
        }
    }

    // Unknown Parts

    void write_unknown_parts();
    void write_unknown_relationships();

    end_part();
}

void xlsx_producer::end_part()
{
    if (current_part_serializer_)
    {
        current_part_serializer_.reset();
    }

    current_part_streambuf_.reset();
}

void xlsx_producer::begin_part(const path &part)
{
    end_part();
    current_part_streambuf_ = archive_->open(part);
    current_part_stream_.rdbuf(current_part_streambuf_.get());
    current_part_serializer_.reset(new xml::serializer(current_part_stream_, part.string()));
}

// Package Parts

void xlsx_producer::write_content_types()
{
    const auto content_types_path = path("[Content_Types].xml");
    begin_part(content_types_path);

    const auto xmlns = "http://schemas.openxmlformats.org/package/2006/content-types";

    write_start_element(xmlns, "Types");
    write_namespace(xmlns, "");

    for (const auto &extension : source_.manifest().extensions_with_default_types())
    {
        write_start_element(xmlns, "Default");
        write_attribute("Extension", extension);
        write_attribute("ContentType", source_.manifest().default_type(extension));
        write_end_element(xmlns, "Default");
    }

    for (const auto &part : source_.manifest().parts_with_overriden_types())
    {
        write_start_element(xmlns, "Override");
        write_attribute("PartName", part.resolve(path("/")).string());
        write_attribute("ContentType", source_.manifest().override_type(part));
        write_end_element(xmlns, "Override");
    }

    write_end_element(xmlns, "Types");
}

void xlsx_producer::write_property(const std::string &name, const variant &value,
    const std::string &ns, bool custom, std::size_t pid)
{
    if (custom)
    {
        write_start_element(ns, "property");
        write_attribute("name", name);
    }
    else
    {
        write_start_element(ns, name);
    }

    switch (value.value_type())
    {
    case variant::type::null:
        {
            break;
        }

    case variant::type::boolean:
        {
            if (custom)
            {
                write_attribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
                write_attribute("pid", pid);
                write_start_element(constants::ns("vt"), "bool");
            }

            write_characters(value.get<bool>() ? "true" : "false");

            if (custom)
            {
                write_end_element(constants::ns("vt"), "bool");
            }

            break;
        }

    case variant::type::i4:
        {
            if (custom)
            {
                write_attribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
                write_attribute("pid", pid);
                write_start_element(constants::ns("vt"), "i4");
            }

            write_characters(value.get<std::int32_t>());

            if (custom)
            {
                write_end_element(constants::ns("vt"), "i4");
            }

            break;
        }

    case variant::type::lpstr:
        {
            if (custom)
            {
                write_attribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
                write_attribute("pid", pid);
                write_start_element(constants::ns("vt"), "lpwstr");
            }

            if (!custom && ns == constants::ns("dcterms") && (name == "created" || name == "modified"))
            {
                write_attribute(xml::qname(constants::ns("xsi"), "type"), "dcterms:W3CDTF");
            }

            write_characters(value.get<std::string>());

            if (custom)
            {
                write_end_element(constants::ns("vt"), "lpwstr");
            }

            break;
        }

    case variant::type::date:
        {
            write_attribute(xml::qname(constants::ns("xsi"), "type"), "dcterms:W3CDTF");
            write_characters(value.get<datetime>().to_iso_string());

            break;
        }

    case variant::type::vector:
        {
            write_start_element(constants::ns("vt"), "vector");

            auto vector = value.get<std::vector<variant>>();
            std::unordered_set<variant::type, scoped_enum_hash<variant::type>> types;

            for (const auto &element : vector)
            {
                types.insert(element.value_type());
            }

            const auto is_mixed = types.size() > 1;
            const auto vector_type = !is_mixed ? to_string(*types.begin()) : "variant";

            write_attribute("size", vector.size());
            write_attribute("baseType", vector_type);

            for (std::size_t i = 0; i < vector.size(); ++i)
            {
                const auto &vector_element = vector.at(i);

                if (is_mixed)
                {
                    write_start_element(constants::ns("vt"), "variant");
                }

                if (vector_element.value_type() == variant::type::lpstr)
                {
                    write_element(constants::ns("vt"), "lpstr", vector_element.get<std::string>());
                }
                else if (vector_element.value_type() == variant::type::i4)
                {
                    write_element(constants::ns("vt"), "i4", vector_element.get<std::int32_t>());
                }

                if (is_mixed)
                {
                    write_end_element(constants::ns("vt"), "variant");
                }
            }

            write_end_element(constants::ns("vt"), "vector");

            break;
        }
    }

    if (custom)
    {
        write_end_element(ns, "property");
    }
    else
    {
        write_end_element(ns, name);
    }
}

void xlsx_producer::write_core_properties(const relationship &/*rel*/)
{
    write_start_element(constants::ns("core-properties"), "coreProperties");

    auto core_properties = source_.core_properties();
    std::unordered_map<std::string, std::string> namespaces;

    write_namespace(constants::ns("core-properties"), "cp");

    for (const auto &prop : core_properties)
    {
        for (const auto &ns : core_property_namespace(prop))
        {
            if (namespaces.count(ns.first) > 0) continue;

            write_namespace(ns.first, ns.second);
            namespaces.emplace(ns);
        }
    }

    for (const auto &prop : core_properties)
    {
        write_property(to_string(prop), source_.core_property(prop),
            core_property_namespace(prop).front().first, false, 0);
    }

    write_end_element(constants::ns("core-properties"), "coreProperties");
}

void xlsx_producer::write_extended_properties(const relationship &/*rel*/)
{
    write_start_element(constants::ns("extended-properties"), "Properties");
    write_namespace(constants::ns("extended-properties"), "");

    if (source_.has_extended_property(extended_property::heading_pairs)
        || source_.has_extended_property(extended_property::titles_of_parts))
    {
        write_namespace(constants::ns("vt"), "vt");
    }

    for (const auto &prop : source_.extended_properties())
    {
        write_property(to_string(prop), source_.extended_property(prop),
            constants::ns("extended-properties"), false, 0);
    }

    write_end_element(constants::ns("extended-properties"), "Properties");
}

void xlsx_producer::write_custom_properties(const relationship &/*rel*/)
{
    write_start_element(constants::ns("custom-properties"), "Properties");
    write_namespace(constants::ns("custom-properties"), "");
    write_namespace(constants::ns("vt"), "vt");

    auto pid = std::size_t(2); // why does this start at 2?

    for (const auto &prop : source_.custom_properties())
    {
        write_property(prop, source_.custom_property(prop),
            constants::ns("custom-properties"), true, pid++);
    }

    write_end_element(constants::ns("custom-properties"), "Properties");
}

// Write SpreadsheetML-Specific Package Parts

void xlsx_producer::write_workbook(const relationship &rel)
{
    std::size_t num_visible = 0;
    bool any_defined_names = false;

    for (auto ws : source_)
    {
        if (!ws.has_page_setup() || ws.page_setup().sheet_state() == sheet_state::visible)
        {
            num_visible++;
        }

        if (ws.has_auto_filter())
        {
            any_defined_names = true;
        }
    }

    if (num_visible == 0)
    {
        throw no_visible_worksheets();
    }

    static const auto &xmlns = constants::ns("workbook");
    static const auto &xmlns_r = constants::ns("r");
    static const auto &xmlns_s = constants::ns("spreadsheetml");

    write_start_element(xmlns, "workbook");
    write_namespace(xmlns, "");
    write_namespace(xmlns_r, "r");

    if (source_.has_file_version())
    {
        write_start_element(xmlns, "fileVersion");

        write_attribute("appName", source_.app_name());
        write_attribute("lastEdited", source_.last_edited());
        write_attribute("lowestEdited", source_.lowest_edited());
        write_attribute("rupBuild", source_.rup_build());

        write_end_element(xmlns, "fileVersion");
    }

    write_start_element(xmlns, "workbookPr");

    if (source_.has_code_name())
    {
        write_attribute("codeName", source_.code_name());
    }

    if (source_.base_date() == calendar::mac_1904)
    {
        write_attribute("date1904", "1");
    }

    write_end_element(xmlns, "workbookPr");

    if (source_.has_view())
    {
        write_start_element(xmlns, "bookViews");
        write_start_element(xmlns, "workbookView");

        const auto &view = source_.view();

        if (view.active_tab.is_set())
        {
            write_attribute("activeTab", view.active_tab.get());
        }

        if (!view.auto_filter_date_grouping)
        {
            write_attribute("autoFilterDateGrouping", write_bool(view.auto_filter_date_grouping));
        }

        if (view.first_sheet.is_set())
        {
            write_attribute("firstSheet", view.first_sheet.get());
        }

        if (view.minimized)
        {
            write_attribute("minimized", write_bool(view.minimized));
        }

        if (!view.show_horizontal_scroll)
        {
            write_attribute("showHorizontalScroll", write_bool(view.show_horizontal_scroll));
        }

        if (!view.show_sheet_tabs)
        {
            write_attribute("showSheetTabs", write_bool(view.show_sheet_tabs));
        }

        if (!view.show_vertical_scroll)
        {
            write_attribute("showVerticalScroll", write_bool(view.show_vertical_scroll));
        }

        if (!view.visible)
        {
            write_attribute("visibility", write_bool(view.visible));
        }

        if (view.x_window.is_set())
        {
            write_attribute("xWindow", view.x_window.get());
        }

        if (view.y_window.is_set())
        {
            write_attribute("yWindow", view.y_window.get());
        }

        if (view.window_width.is_set())
        {
            write_attribute("windowWidth", view.window_width.get());
        }

        if (view.window_height.is_set())
        {
            write_attribute("windowHeight", view.window_height.get());
        }

        if (view.tab_ratio.is_set())
        {
            write_attribute("tabRatio", view.tab_ratio.get());
        }

        write_end_element(xmlns, "workbookView");
        write_end_element(xmlns, "bookViews");
    }

    write_start_element(xmlns, "sheets");

#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wrange-loop-analysis"
    for (const auto ws : source_)
    {
        auto sheet_rel_id = source_.d_->sheet_title_rel_id_map_[ws.title()];
        auto sheet_rel = source_.d_->manifest_.relationship(rel.target().path(), sheet_rel_id);

        write_start_element(xmlns, "sheet");
        write_attribute("name", ws.title());
        write_attribute("sheetId", ws.id());

        if (ws.has_page_setup() && ws.sheet_state() == xlnt::sheet_state::hidden)
        {
            write_attribute("state", "hidden");
        }

        write_attribute(xml::qname(xmlns_r, "id"), sheet_rel_id);
        write_end_element(xmlns, "sheet");
    }
#pragma clang diagnostic pop

    write_end_element(xmlns, "sheets");

    if (any_defined_names)
    {
        write_start_element(xmlns, "definedNames");
        /*
        write_attribute("name", "_xlnm._FilterDatabase");
        write_attribute("hidden", write_bool(true));
        write_attribute("localSheetId", "0");
        write_characters("'" + ws.title() + "'!" +
            range_reference::make_absolute(ws.auto_filter()).to_string());
        */
        write_end_element(xmlns, "definedNames");
    }

    if (source_.has_calculation_properties())
    {
        write_start_element(xmlns, "calcPr");
        write_attribute("calcId", source_.calculation_properties().calc_id);
        //write_attribute("calcMode", "auto");
        //write_attribute("fullCalcOnLoad", "1");
        write_attribute("concurrentCalc", write_bool(source_.calculation_properties().concurrent_calc));
        write_end_element(xmlns, "calcPr");
    }

    if (!source_.named_ranges().empty())
    {
        write_start_element(xmlns, "definedNames");

        for (auto &named_range : source_.named_ranges())
        {
            write_start_element(xmlns_s, "definedName");
            write_namespace(xmlns_s, "s");
            write_attribute("name", named_range.name());
            const auto &target = named_range.targets().front();
            write_characters("'" + target.first.title() + "\'!" + target.second.to_string());
            write_end_element(xmlns_s, "definedName");
        }

        write_end_element(xmlns, "definedNames");
    }

    write_end_element(xmlns, "workbook");

    auto workbook_rels = source_.manifest().relationships(rel.target().path());
    write_relationships(workbook_rels, rel.target().path());

    for (const auto &child_rel : workbook_rels)
    {
        if (child_rel.type() == relationship_type::calculation_chain) continue;

        path archive_path(child_rel.source().path().parent().append(child_rel.target().path()));
        begin_part(archive_path);

        switch (child_rel.type())
        {
        case relationship_type::chartsheet:
            write_chartsheet(child_rel);
            break;

        case relationship_type::connections:
            write_connections(child_rel);
            break;

        case relationship_type::custom_xml_mappings:
            write_custom_xml_mappings(child_rel);
            break;

        case relationship_type::dialogsheet:
            write_dialogsheet(child_rel);
            break;

        case relationship_type::external_workbook_references:
            write_external_workbook_references(child_rel);
            break;

        case relationship_type::pivot_table:
            write_pivot_table(child_rel);
            break;

        case relationship_type::shared_string_table:
            write_shared_string_table(child_rel);
            break;

        case relationship_type::shared_workbook_revision_headers:
            write_shared_workbook_revision_headers(child_rel);
            break;

        case relationship_type::stylesheet:
            write_styles(child_rel);
            break;

        case relationship_type::theme:
            write_theme(child_rel);
            break;

        case relationship_type::volatile_dependencies:
            write_volatile_dependencies(child_rel);
            break;

        case relationship_type::worksheet:
            write_worksheet(child_rel);
            break;

        case relationship_type::calculation_chain:
            break;
        case relationship_type::office_document:
            break;
        case relationship_type::thumbnail:
            break;
        case relationship_type::extended_properties:
            break;
        case relationship_type::core_properties:
            break;
        case relationship_type::hyperlink:
            break;
        case relationship_type::comments:
            break;
        case relationship_type::vml_drawing:
            break;
        case relationship_type::unknown:
            break;
        case relationship_type::custom_properties:
            break;
        case relationship_type::printer_settings:
            break;
        case relationship_type::custom_property:
            break;
        case relationship_type::drawings:
            break;
        case relationship_type::pivot_table_cache_definition:
            break;
        case relationship_type::pivot_table_cache_records:
            break;
        case relationship_type::query_table:
            break;
        case relationship_type::shared_workbook:
            break;
        case relationship_type::revision_log:
            break;
        case relationship_type::shared_workbook_user_data:
            break;
        case relationship_type::single_cell_table_definitions:
            break;
        case relationship_type::table_definition:
            break;
        case relationship_type::image:
            break;
        }
    }
}

// Write Workbook Relationship Target Parts

void xlsx_producer::write_chartsheet(const relationship & /*rel*/)
{
    write_start_element(constants::ns("spreadsheetml"), "chartsheet");
    write_start_element(constants::ns("spreadsheetml"), "chartsheet");
}

void xlsx_producer::write_connections(const relationship & /*rel*/)
{
    write_start_element(constants::ns("spreadsheetml"), "connections");
    write_end_element(constants::ns("spreadsheetml"), "connections");
}

void xlsx_producer::write_custom_xml_mappings(const relationship & /*rel*/)
{
    write_start_element(constants::ns("spreadsheetml"), "MapInfo");
    write_end_element(constants::ns("spreadsheetml"), "MapInfo");
}

void xlsx_producer::write_dialogsheet(const relationship & /*rel*/)
{
    write_start_element(constants::ns("spreadsheetml"), "dialogsheet");
    write_end_element(constants::ns("spreadsheetml"), "dialogsheet");
}

void xlsx_producer::write_external_workbook_references(const relationship & /*rel*/)
{
    write_start_element(constants::ns("spreadsheetml"), "externalLink");
    write_end_element(constants::ns("spreadsheetml"), "externalLink");
}

void xlsx_producer::write_pivot_table(const relationship & /*rel*/)
{
    write_start_element(constants::ns("spreadsheetml"), "pivotTableDefinition");
    write_end_element(constants::ns("spreadsheetml"), "pivotTableDefinition");
}

void xlsx_producer::write_shared_string_table(const relationship & /*rel*/)
{
    static const auto &xmlns = constants::ns("spreadsheetml");

    write_start_element(xmlns, "sst");
    write_namespace(xmlns, "");

    // todo: is there a more elegant way to get this number?
    std::size_t string_count = 0;

#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wrange-loop-analysis"
    for (const auto ws : source_)
    {
        auto dimension = ws.calculate_dimension();
        auto current_cell = dimension.top_left();

        while (current_cell.row() <= dimension.bottom_right().row())
        {
            while (current_cell.column() <= dimension.bottom_right().column())
            {
                if (ws.has_cell(current_cell)
                    && ws.cell(current_cell).data_type() == cell::type::shared_string)
                {
                    ++string_count;
                }

                current_cell.column_index(current_cell.column_index() + 1);
            }

            current_cell.row(current_cell.row() + 1);
            current_cell.column_index(dimension.top_left().column_index());
        }
    }
#pragma clang diagnostic pop

    write_attribute("count", string_count);
    write_attribute("uniqueCount", source_.shared_strings().size());

    auto has_trailing_whitespace = [](const std::string &s)
    {
        return !s.empty() && (s.front() == ' ' || s.back() == ' ');
    };

    for (const auto &string : source_.shared_strings())
    {
        if (string.runs().size() == 1 && !string.runs().at(0).second.is_set())
        {
            write_start_element(xmlns, "si");
            write_start_element(xmlns, "t");
            write_characters(string.plain_text(), has_trailing_whitespace(string.plain_text()));
            write_end_element(xmlns, "t");
            write_end_element(xmlns, "si");

            continue;
        }

        write_start_element(xmlns, "si");

        for (const auto &run : string.runs())
        {
            write_start_element(xmlns, "r");

            if (run.second.is_set())
            {
                write_start_element(xmlns, "rPr");

                if (run.second.get().bold())
                {
                    write_start_element(xmlns, "b");
                    write_end_element(xmlns, "b");
                }

                if (run.second.get().has_size())
                {
                    write_start_element(xmlns, "sz");
                    write_attribute("val", run.second.get().size());
                    write_end_element(xmlns, "sz");
                }

                if (run.second.get().has_color())
                {
                    write_start_element(xmlns, "color");
                    write_color(run.second.get().color());
                    write_end_element(xmlns, "color");
                }

                if (run.second.get().has_name())
                {
                    write_start_element(xmlns, "rFont");
                    write_attribute("val", run.second.get().name());
                    write_end_element(xmlns, "rFont");
                }

                if (run.second.get().has_family())
                {
                    write_start_element(xmlns, "family");
                    write_attribute("val", run.second.get().family());
                    write_end_element(xmlns, "family");
                }

                if (run.second.get().has_scheme())
                {
                    write_start_element(xmlns, "scheme");
                    write_attribute("val", run.second.get().scheme());
                    write_end_element(xmlns, "scheme");
                }

                write_end_element(xmlns, "rPr");
            }

            write_start_element(xmlns, "t");
            write_characters(run.first, has_trailing_whitespace(run.first));
            write_end_element(xmlns, "t");
            write_end_element(xmlns, "r");
        }

        write_end_element(xmlns, "si");
    }

    write_end_element(xmlns, "sst");
}

void xlsx_producer::write_shared_workbook_revision_headers(const relationship & /*rel*/)
{
    write_start_element(constants::ns("spreadsheetml"), "headers");
    write_end_element(constants::ns("spreadsheetml"), "headers");
}

void xlsx_producer::write_shared_workbook(const relationship & /*rel*/)
{
    write_start_element(constants::ns("spreadsheetml"), "revisions");
    write_end_element(constants::ns("spreadsheetml"), "revisions");
}

void xlsx_producer::write_shared_workbook_user_data(const relationship & /*rel*/)
{
    write_start_element(constants::ns("spreadsheetml"), "users");
    write_end_element(constants::ns("spreadsheetml"), "users");
}

void xlsx_producer::write_font(const font &f)
{
	static const auto &xmlns = constants::ns("spreadsheetml");

	write_start_element(xmlns, "font");

	if (f.bold())
	{
		write_start_element(xmlns, "b");
		write_attribute("val", write_bool(true));
		write_end_element(xmlns, "b");
	}

	if (f.italic())
	{
		write_start_element(xmlns, "i");
		write_attribute("val", write_bool(true));
		write_end_element(xmlns, "i");
	}

	if (f.underlined())
	{
		write_start_element(xmlns, "u");
		write_attribute("val", f.underline());
		write_end_element(xmlns, "u");
	}

	if (f.strikethrough())
	{
		write_start_element(xmlns, "strike");
		write_attribute("val", write_bool(true));
		write_end_element(xmlns, "strike");
	}

	write_start_element(xmlns, "sz");
	write_attribute("val", f.size());
	write_end_element(xmlns, "sz");

	if (f.has_color())
	{
		write_start_element(xmlns, "color");
		write_color(f.color());
		write_end_element(xmlns, "color");
	}

	write_start_element(xmlns, "name");
	write_attribute("val", f.name());
	write_end_element(xmlns, "name");

	if (f.has_family())
	{
		write_start_element(xmlns, "family");
		write_attribute("val", f.family());
		write_end_element(xmlns, "family");
	}

	if (f.has_scheme())
	{
		write_start_element(xmlns, "scheme");
		write_attribute("val", f.scheme());
		write_end_element(xmlns, "scheme");
	}

	write_end_element(xmlns, "font");
}

void xlsx_producer::write_fill(const fill &f)
{
	static const auto &xmlns = constants::ns("spreadsheetml");

	write_start_element(xmlns, "fill");

	if (f.type() == xlnt::fill_type::pattern)
	{
		const auto &pattern = f.pattern_fill();

		write_start_element(xmlns, "patternFill");

		write_attribute("patternType", pattern.type());

		if (pattern.foreground().is_set())
		{
			write_start_element(xmlns, "fgColor");
			write_color(pattern.foreground().get());
			write_end_element(xmlns, "fgColor");
		}

		if (pattern.background().is_set())
		{
			write_start_element(xmlns, "bgColor");
			write_color(pattern.background().get());
			write_end_element(xmlns, "bgColor");
		}

		write_end_element(xmlns, "patternFill");
	}
	else if (f.type() == xlnt::fill_type::gradient)
	{
		const auto &gradient = f.gradient_fill();

		write_start_element(xmlns, "gradientFill");
		write_attribute("gradientType", gradient.type());

		if (gradient.degree() != 0.)
		{
			write_attribute("degree", gradient.degree());
		}

		if (gradient.left() != 0.)
		{
			write_attribute("left", gradient.left());
		}

		if (gradient.right() != 0.)
		{
			write_attribute("right", gradient.right());
		}

		if (gradient.top() != 0.)
		{
			write_attribute("top", gradient.top());
		}

		if (gradient.bottom() != 0.)
		{
			write_attribute("bottom", gradient.bottom());
		}

		for (const auto &stop : gradient.stops())
		{
			write_start_element(xmlns, "stop");
			write_attribute("position", stop.first);
			write_start_element(xmlns, "color");
			write_color(stop.second);
			write_end_element(xmlns, "color");
			write_end_element(xmlns, "stop");
		}

		write_end_element(xmlns, "gradientFill");
	}

	write_end_element(xmlns, "fill");
}

void xlsx_producer::write_border(const border &current_border)
{
	static const auto &xmlns = constants::ns("spreadsheetml");

	write_start_element(xmlns, "border");

	if (current_border.diagonal().is_set())
	{
		auto up = current_border.diagonal().get() == diagonal_direction::both
			|| current_border.diagonal().get() == diagonal_direction::up;
		write_attribute("diagonalUp", write_bool(up));

		auto down = current_border.diagonal().get() == diagonal_direction::both
			|| current_border.diagonal().get() == diagonal_direction::down;
		write_attribute("diagonalDown", write_bool(down));
	}

	for (const auto &side : xlnt::border::all_sides())
	{
		if (current_border.side(side).is_set())
		{
			const auto current_side = current_border.side(side).get();

			auto side_name = to_string(side);
			write_start_element(xmlns, side_name);

			if (current_side.style().is_set())
			{
				write_attribute("style", current_side.style().get());
			}

			if (current_side.color().is_set())
			{
				write_start_element(xmlns, "color");
				write_color(current_side.color().get());
				write_end_element(xmlns, "color");
			}

			write_end_element(xmlns, side_name);
		}
	}

	write_end_element(xmlns, "border");
}

void xlsx_producer::write_styles(const relationship & /*rel*/)
{
	static const auto &xmlns = constants::ns("spreadsheetml");

    write_start_element(xmlns, "styleSheet");
    write_namespace(xmlns, "");

    const auto &stylesheet = source_.impl().stylesheet_.get();

    // Number Formats

    if (!stylesheet.number_formats.empty())
    {
        const auto &number_formats = stylesheet.number_formats;

        auto num_custom = std::count_if(number_formats.begin(), number_formats.end(),
            [](const number_format &nf) { return nf.id() >= 164; });

        if (num_custom > 0)
        {
            write_start_element(xmlns, "numFmts");
            write_attribute("count", num_custom);

            for (const auto &num_fmt : number_formats)
            {
                if (num_fmt.id() < 164) continue;
                write_start_element(xmlns, "numFmt");
                write_attribute("numFmtId", num_fmt.id());
                write_attribute("formatCode", num_fmt.format_string());
                write_end_element(xmlns, "numFmt");
            }

            write_end_element(xmlns, "numFmts");
        }
    }

    // Fonts

    if (!stylesheet.fonts.empty())
    {
        const auto &fonts = stylesheet.fonts;

        write_start_element(xmlns, "fonts");
        write_attribute("count", fonts.size());

		for (const auto &current_font : fonts)
		{
			write_font(current_font);
		}

        write_end_element(xmlns, "fonts");
    }

    // Fills

    if (!stylesheet.fills.empty())
    {
        const auto &fills = stylesheet.fills;

        write_start_element(xmlns, "fills");
        write_attribute("count", fills.size());

        for (auto &current_fill : fills)
        {
			write_fill(current_fill);
        }

        write_end_element(xmlns, "fills");
    }

    // Borders

    if (!stylesheet.borders.empty())
    {
        const auto &borders = stylesheet.borders;

        write_start_element(xmlns, "borders");
        write_attribute("count", borders.size());

        for (const auto &current_border : borders)
        {
			write_border(current_border);
        }

        write_end_element(xmlns, "borders");
    }

    // Style XFs
    if (stylesheet.style_impls.size() > 0)
    {
        write_start_element(xmlns, "cellStyleXfs");
        write_attribute("count", stylesheet.style_impls.size());

        for (const auto &current_style_name : stylesheet.style_names)
        {
            const auto &current_style_impl = stylesheet.style_impls.at(current_style_name);

            write_start_element(xmlns, "xf");
            write_attribute("numFmtId", current_style_impl.number_format_id.get());
            write_attribute("fontId", current_style_impl.font_id.get());
            write_attribute("fillId", current_style_impl.fill_id.get());
            write_attribute("borderId", current_style_impl.border_id.get());

            if (current_style_impl.number_format_applied)
            {
                write_attribute("applyNumberFormat", write_bool(true));
            }

            if (current_style_impl.fill_applied)
            {
                write_attribute("applyFill", write_bool(true));
            }

            if (current_style_impl.font_applied)
            {
                write_attribute("applyFont", write_bool(true));
            }

            if (current_style_impl.border_applied)
            {
                write_attribute("applyBorder", write_bool(true));
            }

            if (current_style_impl.alignment_applied)
            {
                write_attribute("applyAlignment", write_bool(true));
            }

            if (current_style_impl.protection_applied)
            {
                write_attribute("applyProtection", write_bool(true));
            }

            if (current_style_impl.pivot_button_)
            {
                write_attribute("pivotButton", write_bool(true));
            }

            if (current_style_impl.quote_prefix_)
            {
                write_attribute("quotePrefix", write_bool(true));
            }

            if (current_style_impl.alignment_id.is_set())
            {
                const auto &current_alignment = stylesheet.alignments[current_style_impl.alignment_id.get()];

                write_start_element(xmlns, "alignment");

                if (current_alignment.vertical().is_set())
                {
                    write_attribute("vertical", current_alignment.vertical().get());
                }

                if (current_alignment.horizontal().is_set())
                {
                    write_attribute("horizontal", current_alignment.horizontal().get());
                }

                if (current_alignment.rotation().is_set())
                {
                    write_attribute("textRotation", current_alignment.rotation().get());
                }

                if (current_alignment.wrap())
                {
                    write_attribute("wrapText", write_bool(current_alignment.wrap()));
                }

                if (current_alignment.indent().is_set())
                {
                    write_attribute("indent", current_alignment.indent().get());
                }

                if (current_alignment.shrink())
                {
                    write_attribute("shrinkToFit", write_bool(current_alignment.shrink()));
                }

                write_end_element(xmlns, "alignment");
            }

            if (current_style_impl.protection_id.is_set())
            {
                const auto &current_protection = stylesheet.protections[current_style_impl.protection_id.get()];

                write_start_element(xmlns, "protection");
                write_attribute("locked", write_bool(current_protection.locked()));
                write_attribute("hidden", write_bool(current_protection.hidden()));
                write_end_element(xmlns, "protection");
            }

            write_end_element(xmlns, "xf");
        }

        write_end_element(xmlns, "cellStyleXfs");
    }

    // Format XFs

    write_start_element(xmlns, "cellXfs");
    write_attribute("count", stylesheet.format_impls.size());

    for (auto &current_format_impl : stylesheet.format_impls)
    {
        write_start_element(xmlns, "xf");

        write_attribute("numFmtId", current_format_impl.number_format_id.get());
        write_attribute("fontId", current_format_impl.font_id.get());

        if (current_format_impl.style.is_set())
        {
            write_attribute("fillId", stylesheet.style_impls.at(current_format_impl.style.get()).fill_id.get());
        }
        else
        {
            write_attribute("fillId", current_format_impl.fill_id.get());
        }

        write_attribute("borderId", current_format_impl.border_id.get());

        if (current_format_impl.number_format_applied)
        {
            write_attribute("applyNumberFormat", write_bool(true));
        }

        if (current_format_impl.fill_applied)
        {
            write_attribute("applyFill", write_bool(true));
        }

        if (current_format_impl.font_applied)
        {
            write_attribute("applyFont", write_bool(true));
        }

        if (current_format_impl.border_applied)
        {
            write_attribute("applyBorder", write_bool(true));
        }

        if (current_format_impl.alignment_applied)
        {
            write_attribute("applyAlignment", write_bool(true));
        }

        if (current_format_impl.protection_applied)
        {
            write_attribute("applyProtection", write_bool(true));
        }

        if (current_format_impl.pivot_button_)
        {
            write_attribute("pivotButton", write_bool(true));
        }

        if (current_format_impl.quote_prefix_)
        {
            write_attribute("quotePrefix", write_bool(true));
        }

        if (current_format_impl.style.is_set())
        {
            write_attribute("xfId", stylesheet.style_index(current_format_impl.style.get()));
        }

        if (current_format_impl.alignment_id.is_set())
        {
            const auto &current_alignment = stylesheet.alignments[current_format_impl.alignment_id.get()];

            write_start_element(xmlns, "alignment");

            if (current_alignment.vertical().is_set())
            {
                write_attribute("vertical", current_alignment.vertical().get());
            }

            if (current_alignment.horizontal().is_set())
            {
                write_attribute("horizontal", current_alignment.horizontal().get());
            }

            if (current_alignment.rotation().is_set())
            {
                write_attribute("textRotation", current_alignment.rotation().get());
            }

            if (current_alignment.wrap())
            {
                write_attribute("wrapText", write_bool(current_alignment.wrap()));
            }

            if (current_alignment.indent().is_set())
            {
                write_attribute("indent", current_alignment.indent().get());
            }

            if (current_alignment.shrink())
            {
                write_attribute("shrinkToFit", write_bool(current_alignment.shrink()));
            }

            write_end_element(xmlns, "alignment");
        }

        if (current_format_impl.protection_id.is_set())
        {
            const auto &current_protection = stylesheet.protections[current_format_impl.protection_id.get()];

            write_start_element(xmlns, "protection");
            write_attribute("locked", write_bool(current_protection.locked()));
            write_attribute("hidden", write_bool(current_protection.hidden()));
            write_end_element(xmlns, "protection");
        }

        write_end_element(xmlns, "xf");
    }

    write_end_element(xmlns, "cellXfs");

    // Styles
    if (stylesheet.style_impls.size() > 0)
    {
        write_start_element(xmlns, "cellStyles");
        write_attribute("count", stylesheet.style_impls.size());
        std::size_t style_index = 0;

        for (auto &current_style_name : stylesheet.style_names)
        {
            const auto &current_style = stylesheet.style_impls.at(current_style_name);

            write_start_element(xmlns, "cellStyle");

            write_attribute("name", current_style.name);
            write_attribute("xfId", style_index++);

            if (current_style.builtin_id.is_set())
            {
                write_attribute("builtinId", current_style.builtin_id.get());
            }

            if (current_style.hidden_style)
            {
                write_attribute("hidden", write_bool(true));
            }

			if (current_style.builtin_id.is_set() && current_style.custom_builtin)
			{
				write_attribute("customBuiltin", write_bool(current_style.custom_builtin));
			}

            write_end_element(xmlns, "cellStyle");
        }

        write_end_element(xmlns, "cellStyles");
    }

	// Conditional Formats
    write_start_element(xmlns, "dxfs");
    write_attribute("count", stylesheet.conditional_format_impls.size());

	for (auto &rule : stylesheet.conditional_format_impls)
	{
		write_start_element(xmlns, "dxf");

		if (rule.border_id.is_set())
		{
			const auto &current_border = stylesheet.borders.at(rule.border_id.get());
			write_border(current_border);
		}

		if (rule.fill_id.is_set())
		{
			const auto &current_fill = stylesheet.fills.at(rule.fill_id.get());
			write_fill(current_fill);
		}

		if (rule.font_id.is_set())
		{
			const auto &current_font = stylesheet.fonts.at(rule.font_id.get());
			write_font(current_font);
		}

		write_end_element(xmlns, "dxf");
	}

    write_end_element(xmlns, "dxfs");

    write_start_element(xmlns, "tableStyles");
    write_attribute("count", "0");
    write_attribute("defaultTableStyle", "TableStyleMedium9");
    write_attribute("defaultPivotStyle", "PivotStyleMedium7");
    write_end_element(xmlns, "tableStyles");

    if (!stylesheet.colors.empty())
    {
        write_start_element(xmlns, "colors");
        write_start_element(xmlns, "indexedColors");

        for (auto &c : stylesheet.colors)
        {
            write_start_element(xmlns, "rgbColor");
            write_attribute("rgb", c.rgb().hex_string());
            write_end_element(xmlns, "rgbColor");
        }

        write_end_element(xmlns, "indexedColors");
        write_end_element(xmlns, "colors");
    }

    write_end_element(xmlns, "styleSheet");
}

void xlsx_producer::write_theme(const relationship &theme_rel)
{
    static const auto &xmlns_a = constants::ns("drawingml");
    static const auto &xmlns_thm15 = constants::ns("thm15");

    write_start_element(xmlns_a, "theme");
    write_namespace(xmlns_a, "a");
    write_attribute("name", "Office Theme");

    write_start_element(xmlns_a, "themeElements");
    write_start_element(xmlns_a, "clrScheme");
    write_attribute("name", "Office");

    struct scheme_element
    {
        std::string name;
        std::string sub_element_name;
        std::string val;
    };

    std::vector<scheme_element> scheme_elements = {
        {"dk1", "sysClr", "windowText"}, {"lt1", "sysClr", "window"}, {"dk2", "srgbClr", "44546A"},
        {"lt2", "srgbClr", "E7E6E6"}, {"accent1", "srgbClr", "5B9BD5"}, {"accent2", "srgbClr", "ED7D31"},
        {"accent3", "srgbClr", "A5A5A5"}, {"accent4", "srgbClr", "FFC000"}, {"accent5", "srgbClr", "4472C4"},
        {"accent6", "srgbClr", "70AD47"}, {"hlink", "srgbClr", "0563C1"}, {"folHlink", "srgbClr", "954F72"},
    };

    for (auto element : scheme_elements)
    {
        write_start_element(xmlns_a, element.name);
        write_start_element(xmlns_a, element.sub_element_name);
        write_attribute("val", element.val);

        if (element.name == "dk1")
        {
            write_attribute("lastClr", "000000");
        }
        else if (element.name == "lt1")
        {
            write_attribute("lastClr", "FFFFFF");
        }

        write_end_element(xmlns_a, element.sub_element_name);
        write_end_element(xmlns_a, element.name);
    }

    write_end_element(xmlns_a, "clrScheme");

    struct font_scheme
    {
        bool typeface;
        std::string script;
        std::string major;
        std::string minor;
    };

    static const auto font_schemes = new std::vector<font_scheme>{{true, "latin", "Calibri Light", "Calibri"},
        {true, "ea", "", ""}, {true, "cs", "", ""}, {false, "Jpan", "Yu Gothic Light", "Yu Gothic"},
        {false, "Hang", "\xeb\xa7\x91\xec\x9d\x80 \xea\xb3\xa0\xeb\x94\x95",
            "\xeb\xa7\x91\xec\x9d\x80 \xea\xb3\xa0\xeb\x94\x95"},
        {false, "Hans", "DengXian Light", "DengXian"},
        {false, "Hant", "\xe6\x96\xb0\xe7\xb4\xb0\xe6\x98\x8e\xe9\xab\x94",
            "\xe6\x96\xb0\xe7\xb4\xb0\xe6\x98\x8e\xe9\xab\x94"},
        {false, "Arab", "Times New Roman", "Arial"}, {false, "Hebr", "Times New Roman", "Arial"},
        {false, "Thai", "Tahoma", "Tahoma"}, {false, "Ethi", "Nyala", "Nyala"}, {false, "Beng", "Vrinda", "Vrinda"},
        {false, "Gujr", "Shruti", "Shruti"}, {false, "Khmr", "MoolBoran", "DaunPenh"},
        {false, "Knda", "Tunga", "Tunga"}, {false, "Guru", "Raavi", "Raavi"}, {false, "Cans", "Euphemia", "Euphemia"},
        {false, "Cher", "Plantagenet Cherokee", "Plantagenet Cherokee"},
        {false, "Yiii", "Microsoft Yi Baiti", "Microsoft Yi Baiti"},
        {false, "Tibt", "Microsoft Himalaya", "Microsoft Himalaya"}, {false, "Thaa", "MV Boli", "MV Boli"},
        {false, "Deva", "Mangal", "Mangal"}, {false, "Telu", "Gautami", "Gautami"}, {false, "Taml", "Latha", "Latha"},
        {false, "Syrc", "Estrangelo Edessa", "Estrangelo Edessa"}, {false, "Orya", "Kalinga", "Kalinga"},
        {false, "Mlym", "Kartika", "Kartika"}, {false, "Laoo", "DokChampa", "DokChampa"},
        {false, "Sinh", "Iskoola Pota", "Iskoola Pota"}, {false, "Mong", "Mongolian Baiti", "Mongolian Baiti"},
        {false, "Viet", "Times New Roman", "Arial"}, {false, "Uigh", "Microsoft Uighur", "Microsoft Uighur"},
        {false, "Geor", "Sylfaen", "Sylfaen"}};

    write_start_element(xmlns_a, "fontScheme");
    write_attribute("name", "Office");

    for (auto major : {true, false})
    {
        write_start_element(xmlns_a, major ? "majorFont" : "minorFont");

        for (const auto &scheme : *font_schemes)
        {
            const auto scheme_value = major ? scheme.major : scheme.minor;

            if (scheme.typeface)
            {
                write_start_element(xmlns_a, scheme.script);
                write_attribute("typeface", scheme_value);

                if (scheme_value == "Calibri Light")
                {
                    write_attribute("panose", "020F0302020204030204");
                }
                else if (scheme_value == "Calibri")
                {
                    write_attribute("panose", "020F0502020204030204");
                }

                write_end_element(xmlns_a, scheme.script);
            }
            else
            {
                write_start_element(xmlns_a, "font");
                write_attribute("script", scheme.script);
                write_attribute("typeface", scheme_value);
                write_end_element(xmlns_a, "font");
            }
        }

        write_end_element(xmlns_a, major ? "majorFont" : "minorFont");
    }

    write_end_element(xmlns_a, "fontScheme");

    write_start_element(xmlns_a, "fmtScheme");
    write_attribute("name", "Office");

    write_start_element(xmlns_a, "fillStyleLst");
    write_start_element(xmlns_a, "solidFill");
    write_start_element(xmlns_a, "schemeClr");
    write_attribute("val", "phClr");
    write_end_element(xmlns_a, "schemeClr");
    write_end_element(xmlns_a, "solidFill");

    write_start_element(xmlns_a, "gradFill");
    write_attribute("rotWithShape", "1");
    write_start_element(xmlns_a, "gsLst");

    write_start_element(xmlns_a, "gs");
    write_attribute("pos", "0");
    write_start_element(xmlns_a, "schemeClr");
    write_attribute("val", "phClr");
    write_start_element(xmlns_a, "lumMod");
    write_attribute("val", "110000");
    write_end_element(xmlns_a, "lumMod");
    write_start_element(xmlns_a, "satMod");
    write_attribute("val", "105000");
    write_end_element(xmlns_a, "satMod");
    write_start_element(xmlns_a, "tint");
    write_attribute("val", "67000");
    write_end_element(xmlns_a, "tint");
    write_end_element(xmlns_a, "schemeClr");
    write_end_element(xmlns_a, "gs");

    write_start_element(xmlns_a, "gs");
    write_attribute("pos", "50000");
    write_start_element(xmlns_a, "schemeClr");
    write_attribute("val", "phClr");
    write_start_element(xmlns_a, "lumMod");
    write_attribute("val", "105000");
    write_end_element(xmlns_a, "lumMod");
    write_start_element(xmlns_a, "satMod");
    write_attribute("val", "103000");
    write_end_element(xmlns_a, "satMod");
    write_start_element(xmlns_a, "tint");
    write_attribute("val", "73000");
    write_end_element(xmlns_a, "tint");
    write_end_element(xmlns_a, "schemeClr");
    write_end_element(xmlns_a, "gs");

    write_start_element(xmlns_a, "gs");
    write_attribute("pos", "100000");
    write_start_element(xmlns_a, "schemeClr");
    write_attribute("val", "phClr");
    write_start_element(xmlns_a, "lumMod");
    write_attribute("val", "105000");
    write_end_element(xmlns_a, "lumMod");
    write_start_element(xmlns_a, "satMod");
    write_attribute("val", "109000");
    write_end_element(xmlns_a, "satMod");
    write_start_element(xmlns_a, "tint");
    write_attribute("val", "81000");
    write_end_element(xmlns_a, "tint");
    write_end_element(xmlns_a, "schemeClr");
    write_end_element(xmlns_a, "gs");

    write_end_element(xmlns_a, "gsLst");

    write_start_element(xmlns_a, "lin");
    write_attribute("ang", "5400000");
    write_attribute("scaled", "0");
    write_end_element(xmlns_a, "lin");

    write_end_element(xmlns_a, "gradFill");

    write_start_element(xmlns_a, "gradFill");
    write_attribute("rotWithShape", "1");
    write_start_element(xmlns_a, "gsLst");

    write_start_element(xmlns_a, "gs");
    write_attribute("pos", "0");
    write_start_element(xmlns_a, "schemeClr");
    write_attribute("val", "phClr");
    write_start_element(xmlns_a, "satMod");
    write_attribute("val", "103000");
    write_end_element(xmlns_a, "satMod");
    write_start_element(xmlns_a, "lumMod");
    write_attribute("val", "102000");
    write_end_element(xmlns_a, "lumMod");
    write_start_element(xmlns_a, "tint");
    write_attribute("val", "94000");
    write_end_element(xmlns_a, "tint");
    write_end_element(xmlns_a, "schemeClr");
    write_end_element(xmlns_a, "gs");

    write_start_element(xmlns_a, "gs");
    write_attribute("pos", "50000");
    write_start_element(xmlns_a, "schemeClr");
    write_attribute("val", "phClr");
    write_start_element(xmlns_a, "satMod");
    write_attribute("val", "110000");
    write_end_element(xmlns_a, "satMod");
    write_start_element(xmlns_a, "lumMod");
    write_attribute("val", "100000");
    write_end_element(xmlns_a, "lumMod");
    write_start_element(xmlns_a, "shade");
    write_attribute("val", "100000");
    write_end_element(xmlns_a, "shade");
    write_end_element(xmlns_a, "schemeClr");
    write_end_element(xmlns_a, "gs");

    write_start_element(xmlns_a, "gs");
    write_attribute("pos", "100000");
    write_start_element(xmlns_a, "schemeClr");
    write_attribute("val", "phClr");
    write_start_element(xmlns_a, "lumMod");
    write_attribute("val", "99000");
    write_end_element(xmlns_a, "lumMod");
    write_start_element(xmlns_a, "satMod");
    write_attribute("val", "120000");
    write_end_element(xmlns_a, "satMod");
    write_start_element(xmlns_a, "shade");
    write_attribute("val", "78000");
    write_end_element(xmlns_a, "shade");
    write_end_element(xmlns_a, "schemeClr");
    write_end_element(xmlns_a, "gs");

    write_end_element(xmlns_a, "gsLst");

    write_start_element(xmlns_a, "lin");
    write_attribute("ang", "5400000");
    write_attribute("scaled", "0");
    write_end_element(xmlns_a, "lin");

    write_end_element(xmlns_a, "gradFill");
    write_end_element(xmlns_a, "fillStyleLst");

    write_start_element(xmlns_a, "lnStyleLst");

    write_start_element(xmlns_a, "ln");
    write_attribute("w", "6350");
    write_attribute("cap", "flat");
    write_attribute("cmpd", "sng");
    write_attribute("algn", "ctr");
    write_start_element(xmlns_a, "solidFill");
    write_start_element(xmlns_a, "schemeClr");
    write_attribute("val", "phClr");
    write_end_element(xmlns_a, "schemeClr");
    write_end_element(xmlns_a, "solidFill");
    write_start_element(xmlns_a, "prstDash");
    write_attribute("val", "solid");
    write_end_element(xmlns_a, "prstDash");
    write_start_element(xmlns_a, "miter");
    write_attribute("lim", "800000");
    write_end_element(xmlns_a, "miter");
    write_end_element(xmlns_a, "ln");

    write_start_element(xmlns_a, "ln");
    write_attribute("w", "12700");
    write_attribute("cap", "flat");
    write_attribute("cmpd", "sng");
    write_attribute("algn", "ctr");
    write_start_element(xmlns_a, "solidFill");
    write_start_element(xmlns_a, "schemeClr");
    write_attribute("val", "phClr");
    write_end_element(xmlns_a, "schemeClr");
    write_end_element(xmlns_a, "solidFill");
    write_start_element(xmlns_a, "prstDash");
    write_attribute("val", "solid");
    write_end_element(xmlns_a, "prstDash");
    write_start_element(xmlns_a, "miter");
    write_attribute("lim", "800000");
    write_end_element(xmlns_a, "miter");
    write_end_element(xmlns_a, "ln");

    write_start_element(xmlns_a, "ln");
    write_attribute("w", "19050");
    write_attribute("cap", "flat");
    write_attribute("cmpd", "sng");
    write_attribute("algn", "ctr");
    write_start_element(xmlns_a, "solidFill");
    write_start_element(xmlns_a, "schemeClr");
    write_attribute("val", "phClr");
    write_end_element(xmlns_a, "schemeClr");
    write_end_element(xmlns_a, "solidFill");
    write_start_element(xmlns_a, "prstDash");
    write_attribute("val", "solid");
    write_end_element(xmlns_a, "prstDash");
    write_start_element(xmlns_a, "miter");
    write_attribute("lim", "800000");
    write_end_element(xmlns_a, "miter");
    write_end_element(xmlns_a, "ln");

    write_end_element(xmlns_a, "lnStyleLst");

    write_start_element(xmlns_a, "effectStyleLst");

    write_start_element(xmlns_a, "effectStyle");
    write_element(xmlns_a, "effectLst", "");
    write_end_element(xmlns_a, "effectStyle");

    write_start_element(xmlns_a, "effectStyle");
    write_element(xmlns_a, "effectLst", "");
    write_end_element(xmlns_a, "effectStyle");

    write_start_element(xmlns_a, "effectStyle");
    write_start_element(xmlns_a, "effectLst");
    write_start_element(xmlns_a, "outerShdw");
    write_attribute("blurRad", "57150");
    write_attribute("dist", "19050");
    write_attribute("dir", "5400000");
    write_attribute("algn", "ctr");
    write_attribute("rotWithShape", "0");
    write_start_element(xmlns_a, "srgbClr");
    write_attribute("val", "000000");
    write_start_element(xmlns_a, "alpha");
    write_attribute("val", "63000");
    write_end_element(xmlns_a, "alpha");
    write_end_element(xmlns_a, "srgbClr");
    write_end_element(xmlns_a, "outerShdw");
    write_end_element(xmlns_a, "effectLst");
    write_end_element(xmlns_a, "effectStyle");

    write_end_element(xmlns_a, "effectStyleLst");

    write_start_element(xmlns_a, "bgFillStyleLst");

    write_start_element(xmlns_a, "solidFill");
    write_start_element(xmlns_a, "schemeClr");
    write_attribute("val", "phClr");
    write_end_element(xmlns_a, "schemeClr");
    write_end_element(xmlns_a, "solidFill");

    write_start_element(xmlns_a, "solidFill");
    write_start_element(xmlns_a, "schemeClr");
    write_attribute("val", "phClr");
    write_start_element(xmlns_a, "tint");
    write_attribute("val", "95000");
    write_end_element(xmlns_a, "tint");
    write_start_element(xmlns_a, "satMod");
    write_attribute("val", "170000");
    write_end_element(xmlns_a, "satMod");
    write_end_element(xmlns_a, "schemeClr");
    write_end_element(xmlns_a, "solidFill");

    write_start_element(xmlns_a, "gradFill");
    write_attribute("rotWithShape", "1");
    write_start_element(xmlns_a, "gsLst");

    write_start_element(xmlns_a, "gs");
    write_attribute("pos", "0");
    write_start_element(xmlns_a, "schemeClr");
    write_attribute("val", "phClr");
    write_start_element(xmlns_a, "tint");
    write_attribute("val", "93000");
    write_end_element(xmlns_a, "tint");
    write_start_element(xmlns_a, "satMod");
    write_attribute("val", "150000");
    write_end_element(xmlns_a, "satMod");
    write_start_element(xmlns_a, "shade");
    write_attribute("val", "98000");
    write_end_element(xmlns_a, "shade");
    write_start_element(xmlns_a, "lumMod");
    write_attribute("val", "102000");
    write_end_element(xmlns_a, "lumMod");
    write_end_element(xmlns_a, "schemeClr");
    write_end_element(xmlns_a, "gs");

    write_start_element(xmlns_a, "gs");
    write_attribute("pos", "50000");
    write_start_element(xmlns_a, "schemeClr");
    write_attribute("val", "phClr");
    write_start_element(xmlns_a, "tint");
    write_attribute("val", "98000");
    write_end_element(xmlns_a, "tint");
    write_start_element(xmlns_a, "satMod");
    write_attribute("val", "130000");
    write_end_element(xmlns_a, "satMod");
    write_start_element(xmlns_a, "shade");
    write_attribute("val", "90000");
    write_end_element(xmlns_a, "shade");
    write_start_element(xmlns_a, "lumMod");
    write_attribute("val", "103000");
    write_end_element(xmlns_a, "lumMod");
    write_end_element(xmlns_a, "schemeClr");
    write_end_element(xmlns_a, "gs");

    write_start_element(xmlns_a, "gs");
    write_attribute("pos", "100000");
    write_start_element(xmlns_a, "schemeClr");
    write_attribute("val", "phClr");
    write_start_element(xmlns_a, "shade");
    write_attribute("val", "63000");
    write_end_element(xmlns_a, "shade");
    write_start_element(xmlns_a, "satMod");
    write_attribute("val", "120000");
    write_end_element(xmlns_a, "satMod");
    write_end_element(xmlns_a, "schemeClr");
    write_end_element(xmlns_a, "gs");

    write_end_element(xmlns_a, "gsLst");

    write_start_element(xmlns_a, "lin");
    write_attribute("ang", "5400000");
    write_attribute("scaled", "0");
    write_end_element(xmlns_a, "lin");

    write_end_element(xmlns_a, "gradFill");

    write_end_element(xmlns_a, "bgFillStyleLst");
    write_end_element(xmlns_a, "fmtScheme");
    write_end_element(xmlns_a, "themeElements");

    write_element(xmlns_a, "objectDefaults", "");
    write_element(xmlns_a, "extraClrSchemeLst", "");

    write_start_element(xmlns_a, "extLst");
    write_start_element(xmlns_a, "ext");
    write_attribute("uri", "{05A4C25C-085E-4340-85A3-A5531E510DB2}");
    write_start_element(xmlns_thm15, "themeFamily");
    write_namespace(xmlns_thm15, "thm15");
    write_attribute("name", "Office Theme");
    write_attribute("id", "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}");
    write_attribute("vid", "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}");
    write_end_element(xmlns_thm15, "themeFamily");
    write_end_element(xmlns_a, "ext");
    write_end_element(xmlns_a, "extLst");

    write_end_element(xmlns_a, "theme");

    const auto workbook_rel = source_.manifest().relationship(path("/"), relationship_type::office_document);
    const auto theme_part = source_.manifest().canonicalize({workbook_rel, theme_rel});
    const auto theme_rels = source_.manifest().relationships(theme_part);

    if (!theme_rels.empty())
    {
        write_relationships(theme_rels, theme_part);

        for (auto rel : theme_rels)
        {
            if (rel.type() == relationship_type::image)
            {
                const auto image_path = source_.manifest().canonicalize({workbook_rel, theme_rel, rel});
                write_image(image_path);
            }
        }
    }
}

void xlsx_producer::write_volatile_dependencies(const relationship & /*rel*/)
{
    write_start_element(constants::ns("spreadsheetml"), "volTypes");
    write_end_element(constants::ns("spreadsheetml"), "volTypes");
}

void xlsx_producer::write_worksheet(const relationship &rel)
{
    static const auto &xmlns = constants::ns("spreadsheetml");
    static const auto &xmlns_r = constants::ns("r");

    auto worksheet_part = rel.source().path().parent().append(rel.target().path());
    auto worksheet_rels = source_.manifest().relationships(worksheet_part);

    auto title = std::find_if(source_.d_->sheet_title_rel_id_map_.begin(), source_.d_->sheet_title_rel_id_map_.end(),
        [&](const std::pair<std::string, std::string> &p) {
            return p.second == rel.id();
        })->first;

    auto ws = source_.sheet_by_title(title);

    write_start_element(xmlns, "worksheet");
    write_namespace(xmlns, "");
    write_namespace(xmlns_r, "r");

    if (ws.has_page_setup())
    {
        write_start_element(xmlns, "sheetPr");

        write_start_element(xmlns, "outlinePr");
        write_attribute("summaryBelow", "1");
        write_attribute("summaryRight", "1");
        write_end_element(xmlns, "outlinePr");

        write_start_element(xmlns, "pageSetUpPr");
        write_attribute("fitToPage", write_bool(ws.page_setup().fit_to_page()));
        write_end_element(xmlns, "pageSetUpPr");

        write_end_element(xmlns, "sheetPr");
    }

    write_start_element(xmlns, "dimension");
    const auto dimension = ws.calculate_dimension();
    write_attribute("ref", dimension.is_single_cell()
        ? dimension.top_left().to_string()
        : dimension.to_string());
    write_end_element(xmlns, "dimension");

    if (ws.has_view())
    {
        write_start_element(xmlns, "sheetViews");
        write_start_element(xmlns, "sheetView");

        const auto view = ws.view();

        write_attribute("tabSelected", write_bool(view.id() == 0));
        write_attribute("workbookViewId", view.id());

        if (view.type() != sheet_view_type::normal)
        {
            write_attribute(
                "view", view.type() == sheet_view_type::page_break_preview ? "pageBreakPreview" : "pageLayout");
        }

        if (view.has_pane())
        {
            const auto &current_pane = view.pane();
            write_start_element(xmlns, "pane"); // CT_Pane

            if (current_pane.top_left_cell.is_set())
            {
                write_attribute("topLeftCell", current_pane.top_left_cell.get().to_string());
            }

            write_attribute("xSplit", current_pane.x_split.index);
            write_attribute("ySplit", current_pane.y_split);

            if (current_pane.active_pane != pane_corner::top_left)
            {
                write_attribute("activePane", current_pane.active_pane);
            }

            if (current_pane.state != pane_state::split)
            {
                write_attribute("state", current_pane.state);
            }

            write_end_element(xmlns, "pane");
        }

        for (const auto &current_selection : view.selections())
        {
            write_start_element(xmlns, "selection"); // CT_Selection

            if (current_selection.has_active_cell())
            {
                write_attribute("activeCell", current_selection.active_cell().to_string());
                write_attribute("sqref", current_selection.active_cell().to_string());
            }
            /*
                        if (current_selection.sqref() != "A1:A1")
                        {
                            write_attribute("sqref", current_selection.sqref().to_string());
                        }
            */
            if (current_selection.pane() != pane_corner::top_left)
            {
                write_attribute("pane", current_selection.pane());
            }

            write_end_element(xmlns, "selection");
        }

        write_end_element(xmlns, "sheetView");
        write_end_element(xmlns, "sheetViews");
    }

    write_start_element(xmlns, "sheetFormatPr");
    write_attribute("baseColWidth", "10");
    write_attribute("defaultRowHeight", "16");
    write_end_element(xmlns, "sheetFormatPr");

    bool has_column_properties = false;

    for (auto column = ws.lowest_column(); column <= ws.highest_column(); column++)
    {
        if (ws.has_column_properties(column))
        {
            has_column_properties = true;
            break;
        }
    }

    if (has_column_properties)
    {
        write_start_element(xmlns, "cols");

        for (auto column = ws.lowest_column_or_props(); column <= ws.highest_column_or_props(); column++)
        {
            if (!ws.has_column_properties(column)) continue;

            const auto &props = ws.column_properties(column);

            write_start_element(xmlns, "col");
            write_attribute("min", column.index);
            write_attribute("max", column.index);

            if (props.width.is_set())
            {
                write_attribute("width", (props.width.get() * 7 + 5) / 7);
            }

            if (props.custom_width)
            {
                write_attribute("customWidth", write_bool(true));
            }

            if (props.style.is_set())
            {
                write_attribute("style", props.style.get());
            }

            if (props.hidden)
            {
                write_attribute("hidden", write_bool(true));
            }

            write_end_element(xmlns, "col");
        }

        write_end_element(xmlns, "cols");
    }

    const auto hyperlink_rels = source_.manifest().relationships(worksheet_part, relationship_type::hyperlink);
    std::unordered_map<std::string, std::string> reverse_hyperlink_references;

    for (auto hyperlink_rel : hyperlink_rels)
    {
        reverse_hyperlink_references[hyperlink_rel.target().path().string()] = rel.id();
    }

    std::unordered_map<std::string, std::string> hyperlink_references;
    std::vector<cell_reference> cells_with_comments;

    write_start_element(xmlns, "sheetData");

    for (auto row = ws.lowest_row_or_props(); row <= ws.highest_row_or_props(); ++row)
    {
        auto first_column = constants::max_column();
        auto last_column = constants::min_column();

        bool any_non_null = false;

        for (auto column = dimension.top_left().column(); column <= dimension.bottom_right().column(); ++column)
        {
            if (!ws.has_cell(cell_reference(column, row))) continue;

            auto cell = ws.cell(cell_reference(column, row));

            first_column = std::min(first_column, cell.column());
            last_column = std::max(last_column, cell.column());

            if (!cell.garbage_collectible())
            {
                any_non_null = true;
            }
        }

        if (!any_non_null && !ws.has_row_properties(row)) continue;

        write_start_element(xmlns, "row");
        write_attribute("r", row);

        if (any_non_null)
        {
            auto span_string = std::to_string(first_column.index) + ":" + std::to_string(last_column.index);
            write_attribute("spans", span_string);
        }

        if (ws.has_row_properties(row))
        {
            const auto &props = ws.row_properties(row);

            if (props.custom_height)
            {
                write_attribute("customHeight", write_bool(true));
            }

            if (props.height.is_set())
            {
                auto height = props.height.get();

                if (std::fabs(height - std::floor(height)) == 0.0)
                {
                    write_attribute("ht", std::to_string(static_cast<int>(height)) + ".0");
                }
                else
                {
                    write_attribute("ht", height);
                }
            }

            if (props.hidden)
            {
                write_attribute("hidden", write_bool(true));
            }
        }

        if (any_non_null)
        {
            for (auto column = dimension.top_left().column(); column <= dimension.bottom_right().column(); ++column)
            {
                if (!ws.has_cell(cell_reference(column, row))) continue;

                auto cell = ws.cell(cell_reference(column, row));

                if (cell.garbage_collectible()) continue;

                // record data about the cell needed later

                if (cell.has_comment())
                {
                    cells_with_comments.push_back(cell.reference());
                }

                if (cell.has_hyperlink())
                {
                    hyperlink_references[cell.reference().to_string()] = reverse_hyperlink_references[cell.hyperlink()];
                }

                write_start_element(xmlns, "c");

                // begin cell attributes

                write_attribute("r", cell.reference().to_string());

                if (cell.has_format())
                {
                    write_attribute("s", cell.format().d_->id);
                }

                switch (cell.data_type())
                {
                case cell::type::empty:
                    break;

                case cell::type::boolean:
                    write_attribute("t", "b");
                    break;

                case cell::type::date:
                    write_attribute("t", "d");
                    break;

                case cell::type::error:
                    write_attribute("t", "e");
                    break;

                case cell::type::inline_string:
                    write_attribute("t", "inlineStr");
                    break;

                case cell::type::number:
                    write_attribute("t", "n");
                    break;

                case cell::type::shared_string:
                    write_attribute("t", "s");
                    break;

                case cell::type::formula_string:
                    write_attribute("t", "str");
                    break;
                }

                //write_attribute("cm", "");
                //write_attribute("vm", "");
                //write_attribute("ph", "");

                // begin child elements

                if (cell.has_formula())
                {
                    write_element(xmlns, "f", cell.formula());
                }

                switch (cell.data_type())
                {
                case cell::type::empty:
                    break;

                case cell::type::boolean:
                    write_element(xmlns, "v", write_bool(cell.value<bool>()));
                    break;

                case cell::type::date:
                    write_element(xmlns, "v", cell.value<std::string>());
                    break;

                case cell::type::error:
                    write_element(xmlns, "v", cell.value<std::string>());
                    break;

                case cell::type::inline_string:
                    write_start_element(xmlns, "is");
                    // TODO: make a write_rich_text method and use that here
                    write_element(xmlns, "t", cell.value<std::string>());
                    write_end_element(xmlns, "is");
                    break;

                case cell::type::number:
                    write_start_element(xmlns, "v");

                    if (is_integral(cell.value<double>()))
                    {
                        write_characters(static_cast<std::int64_t>(cell.value<double>()));
                    }
                    else
                    {
                        std::stringstream ss;
                        ss.precision(20);
                        ss << cell.value<double>();
                        write_characters(ss.str());
                    }

                    write_end_element(xmlns, "v");
                    break;

                case cell::type::shared_string:
                    write_element(xmlns, "v", static_cast<std::size_t>(cell.d_->value_numeric_));
                    break;

                case cell::type::formula_string:
                    write_element(xmlns, "v", cell.value<std::string>());
                    break;
                }

                write_end_element(xmlns, "c");
            }
        }

        write_end_element(xmlns, "row");
    }

    write_end_element(xmlns, "sheetData");

    if (ws.has_auto_filter())
    {
        write_start_element(xmlns, "autoFilter");
        write_attribute("ref", ws.auto_filter().to_string());
        write_end_element(xmlns, "autoFilter");
    }

    if (!ws.merged_ranges().empty())
    {
        write_start_element(xmlns, "mergeCells");
        write_attribute("count", ws.merged_ranges().size());

        for (auto merged_range : ws.merged_ranges())
        {
            write_start_element(xmlns, "mergeCell");
            write_attribute("ref", merged_range.to_string());
            write_end_element(xmlns, "mergeCell");
        }

        write_end_element(xmlns, "mergeCells");
    }

	if (source_.impl().stylesheet_.is_set())
	{
		const auto &stylesheet = source_.impl().stylesheet_.get();
		const auto &cf_impls = stylesheet.conditional_format_impls;

		std::unordered_map<std::string, std::vector<const conditional_format_impl *>> range_map;

		for (auto &cf : cf_impls)
		{
			if (cf.target_sheet != ws.d_) continue;

			if (range_map.find(cf.target_range.to_string()) == range_map.end())
			{
				range_map[cf.target_range.to_string()] = {};
			}

			range_map[cf.target_range.to_string()].push_back(&cf);
		}

		for (const auto &range_rules_pair : range_map)
		{
			write_start_element(xmlns, "conditionalFormatting");
			write_attribute("sqref", range_rules_pair.first);

			std::size_t i = 1;

			for (auto rule : range_rules_pair.second)
			{
				write_start_element(xmlns, "cfRule");
				write_attribute("type", "containsText");
				write_attribute("operator", "containsText");
				write_attribute("dxfId", rule->differential_format_id);
				write_attribute("priority", i++);
				write_attribute("text", rule->when.text_comparand_);
				//TODO: what does this formula mean and why is it necessary?
				write_element(xmlns, "formula", "NOT(ISERROR(SEARCH(\"" + rule->when.text_comparand_ + "\",A1)))");
				write_end_element(xmlns, "cfRule");
			}

			write_end_element(xmlns, "conditionalFormatting");
		}
	}

    if (!hyperlink_rels.empty())
    {
        write_start_element(xmlns, "hyperlinks");

        for (const auto &hyperlink : hyperlink_references)
        {
            write_start_element(xmlns, "hyperlink");
            write_attribute("ref", hyperlink.first);
            write_attribute(xml::qname(xmlns_r, "id"), hyperlink.second);
            write_end_element(xmlns, "hyperlink");
        }

        write_end_element(xmlns, "hyperlinks");
    }

    if (ws.has_page_setup())
    {
        write_start_element(xmlns, "printOptions");
        write_attribute("horizontalCentered", write_bool(ws.page_setup().horizontal_centered()));
        write_attribute("verticalCentered", write_bool(ws.page_setup().vertical_centered()));
        write_end_element(xmlns, "printOptions");
    }

    if (ws.has_page_margins())
    {
        write_start_element(xmlns, "pageMargins");

        // TODO: there must be a better way to do this
        auto remove_trailing_zeros = [](const std::string &n) -> std::string {
            auto decimal = n.find('.');

            if (decimal == std::string::npos) return n;

            auto index = n.size() - 1;

            while (index >= decimal && n[index] == '0')
            {
                index--;
            }

            if (index == decimal)
            {
                return n.substr(0, decimal);
            }

            return n.substr(0, index + 1);
        };

        write_attribute("left", remove_trailing_zeros(std::to_string(ws.page_margins().left())));
        write_attribute("right", remove_trailing_zeros(std::to_string(ws.page_margins().right())));
        write_attribute("top", remove_trailing_zeros(std::to_string(ws.page_margins().top())));
        write_attribute("bottom", remove_trailing_zeros(std::to_string(ws.page_margins().bottom())));
        write_attribute("header", remove_trailing_zeros(std::to_string(ws.page_margins().header())));
        write_attribute("footer", remove_trailing_zeros(std::to_string(ws.page_margins().footer())));

        write_end_element(xmlns, "pageMargins");
    }

    if (ws.has_page_setup())
    {
        write_start_element(xmlns, "pageSetup");
        write_attribute(
            "orientation", ws.page_setup().orientation() == xlnt::orientation::landscape ? "landscape" : "portrait");
        write_attribute("paperSize", static_cast<std::size_t>(ws.page_setup().paper_size()));
        write_attribute("fitToHeight", write_bool(ws.page_setup().fit_to_height()));
        write_attribute("fitToWidth", write_bool(ws.page_setup().fit_to_width()));
        write_end_element(xmlns, "pageSetup");
    }

    if (ws.has_header_footer())
    {
        const auto hf = ws.header_footer();

        write_start_element(xmlns, "headerFooter");

        auto odd_header = std::string();
        auto odd_footer = std::string();
        auto even_header = std::string();
        auto even_footer = std::string();
        auto first_header = std::string();
        auto first_footer = std::string();

        const auto locations =
        {
            header_footer::location::left,
            header_footer::location::center,
            header_footer::location::right
        };

        using xlnt::detail::encode_header_footer;

        for (auto location : locations)
        {
            if (hf.different_odd_even())
            {
                if (hf.has_odd_even_header(location))
                {
                    odd_header.append(encode_header_footer(hf.odd_header(location), location));
                    even_header.append(encode_header_footer(hf.even_header(location), location));
                }

                if (hf.has_odd_even_footer(location))
                {
                    odd_footer.append(encode_header_footer(hf.odd_footer(location), location));
                    even_footer.append(encode_header_footer(hf.even_footer(location), location));
                }
            }
            else
            {
                if (hf.has_header(location))
                {
                    odd_header.append(encode_header_footer(hf.header(location), location));
                }

                if (hf.has_footer(location))
                {
                    odd_footer.append(encode_header_footer(hf.footer(location), location));
                }
            }

            if (hf.different_first())
            {
                if (hf.has_first_page_header(location))
                {
                    first_header.append(encode_header_footer(hf.first_page_header(location), location));
                }

                if (hf.has_first_page_footer(location))
                {
                    first_footer.append(encode_header_footer(hf.first_page_footer(location), location));
                }
            }
        }

        if (!odd_header.empty())
        {
            write_element(xmlns, "oddHeader", odd_header);
        }

        if (!odd_footer.empty())
        {
            write_element(xmlns, "oddFooter", odd_footer);
        }

        if (!even_header.empty())
        {
            write_element(xmlns, "evenHeader", even_header);
        }

        if (!even_footer.empty())
        {
            write_element(xmlns, "evenFooter", even_footer);
        }

        if (!first_header.empty())
        {
            write_element(xmlns, "firstHeader", first_header);
        }

        if (!first_footer.empty())
        {
            write_element(xmlns, "firstFooter", first_footer);
        }

        write_end_element(xmlns, "headerFooter");
    }

    if (!ws.page_break_rows().empty())
    {
        write_start_element(xmlns, "rowBreaks");

        write_attribute("count", ws.page_break_rows().size());
        write_attribute("manualBreakCount", ws.page_break_rows().size());

        for (auto break_id : ws.page_break_rows())
        {
            write_start_element(xmlns, "brk");
            write_attribute("id", break_id);
            write_attribute("max", 16383);
            write_attribute("man", 1);
            write_end_element(xmlns, "brk");
        }

        write_end_element(xmlns, "rowBreaks");
    }

    if (!ws.page_break_columns().empty())
    {
        write_start_element(xmlns, "colBreaks");

        write_attribute("count", ws.page_break_columns().size());
        write_attribute("manualBreakCount", ws.page_break_columns().size());

        for (auto break_id : ws.page_break_columns())
        {
            write_start_element(xmlns, "brk");
            write_attribute("id", break_id.index);
            write_attribute("max", 1048575);
            write_attribute("man", 1);
            write_end_element(xmlns, "brk");
        }

        write_end_element(xmlns, "colBreaks");
    }

    if (!worksheet_rels.empty())
    {
        for (const auto &child_rel : worksheet_rels)
        {
            if (child_rel.type() == xlnt::relationship_type::vml_drawing)
            {
                write_start_element(xmlns, "legacyDrawing");
                write_attribute(xml::qname(xmlns_r, "id"), child_rel.id());
                write_end_element(xmlns, "legacyDrawing");

                // todo: there's only one of these per sheet, right?
                break;
            }
        }
    }

    write_end_element(xmlns, "worksheet");

    if (!worksheet_rels.empty())
    {
        write_relationships(worksheet_rels, worksheet_part);

        for (const auto &child_rel : worksheet_rels)
        {
            if (child_rel.target_mode() == target_mode::external) continue;

            // todo: this is ugly
            path archive_path(worksheet_part.parent().append(child_rel.target().path()));
            auto split_part_path = archive_path.split();
            auto part_path_iter = split_part_path.begin();

            while (part_path_iter != split_part_path.end())
            {
                if (*part_path_iter == "..")
                {
                    part_path_iter = split_part_path.erase(part_path_iter - 1, part_path_iter + 1);
                    continue;
                }

                ++part_path_iter;
            }

            archive_path = std::accumulate(split_part_path.begin(), split_part_path.end(), path(""),
                [](const path &a, const std::string &b) { return a.append(b); });

            begin_part(archive_path);

            if (child_rel.type() == relationship_type::comments)
            {
                write_comments(child_rel, ws, cells_with_comments);
            }
            else if (child_rel.type() == relationship_type::vml_drawing)
            {
                write_vml_drawings(child_rel, ws, cells_with_comments);
            }
        }
    }
}

// Sheet Relationship Target Parts

void xlsx_producer::write_comments(const relationship & /*rel*/, worksheet ws, const std::vector<cell_reference> &cells)
{
    static const auto &xmlns = constants::ns("spreadsheetml");

    write_start_element(xmlns, "comments");
    write_namespace(xmlns, "");

    if (!cells.empty())
    {
        std::unordered_map<std::string, std::size_t> authors;

        for (auto cell_ref : cells)
        {
            auto cell = ws.cell(cell_ref);
            auto author = cell.comment().author();

            if (authors.find(author) == authors.end())
            {
                auto author_index = authors.size();
                authors[author] = author_index;
            }
        }

        write_start_element(xmlns, "authors");

        for (const auto &author : authors)
        {
            write_start_element(xmlns, "author");
            write_characters(author.first);
            write_end_element(xmlns, "author");
        }

        write_end_element(xmlns, "authors");
        write_start_element(xmlns, "commentList");

        for (const auto &cell_ref : cells)
        {
            write_start_element(xmlns, "comment");

            auto cell = ws.cell(cell_ref);
            auto cell_comment = cell.comment();

            write_attribute("ref", cell_ref.to_string());
            auto author_id = authors.at(cell_comment.author());
            write_attribute("authorId", author_id);
            write_start_element(xmlns, "text");

            for (const auto &run : cell_comment.text().runs())
            {
                write_start_element(xmlns, "r");

                if (run.second.is_set())
                {
                    write_start_element(xmlns, "rPr");

                    if (run.second.get().bold())
                    {
                        write_start_element(xmlns, "b");
                        write_end_element(xmlns, "b");
                    }

                    if (run.second.get().has_size())
                    {
                        write_start_element(xmlns, "sz");
                        write_attribute("val", run.second.get().size());
                        write_end_element(xmlns, "sz");
                    }

                    if (run.second.get().has_color())
                    {
                        write_start_element(xmlns, "color");
                        write_color(run.second.get().color());
                        write_end_element(xmlns, "color");
                    }

                    if (run.second.get().has_name())
                    {
                        write_start_element(xmlns, "rFont");
                        write_attribute("val", run.second.get().name());
                        write_end_element(xmlns, "rFont");
                    }

                    if (run.second.get().has_family())
                    {
                        write_start_element(xmlns, "family");
                        write_attribute("val", run.second.get().family());
                        write_end_element(xmlns, "family");
                    }

                    if (run.second.get().has_scheme())
                    {
                        write_start_element(xmlns, "scheme");
                        write_attribute("val", run.second.get().scheme());
                        write_end_element(xmlns, "scheme");
                    }

                    write_end_element(xmlns, "rPr");
                }

                write_element(xmlns, "t", run.first);
                write_end_element(xmlns, "r");
            }

            write_end_element(xmlns, "text");
            write_end_element(xmlns, "comment");
        }

        write_end_element(xmlns, "commentList");
    }

    write_end_element(xmlns, "comments");
}

void xlsx_producer::write_vml_drawings(const relationship &rel, worksheet ws, const std::vector<cell_reference> &cells)
{
    static const auto &xmlns_mv = std::string("http://macVmlSchemaUri");
    static const auto &xmlns_o = std::string("urn:schemas-microsoft-com:office:office");
    static const auto &xmlns_v = std::string("urn:schemas-microsoft-com:vml");
    static const auto &xmlns_x = std::string("urn:schemas-microsoft-com:office:excel");

    write_start_element("xml");
    write_namespace(xmlns_v, "v");
    write_namespace(xmlns_o, "o");
    write_namespace(xmlns_x, "x");
    write_namespace(xmlns_mv, "mv");

    write_start_element(xmlns_o, "shapelayout");
    write_attribute(xml::qname(xmlns_v, "ext"), "edit");
    write_start_element(xmlns_o, "idmap");
    write_attribute(xml::qname(xmlns_v, "ext"), "edit");

    auto filename = rel.target().path().split_extension().first;
    auto index_pos = filename.size() - 1;

    while (filename[index_pos] >= '0' && filename[index_pos] <= '9')
    {
        index_pos--;
    }

    auto file_index = std::stoull(filename.substr(index_pos + 1));

    write_attribute("data", file_index);
    write_end_element(xmlns_o, "idmap");
    write_end_element(xmlns_o, "shapelayout");

    write_start_element(xmlns_v, "shapetype");
    write_attribute("id", "_x0000_t202");
    write_attribute("coordsize", "21600,21600");
    write_attribute(xml::qname(xmlns_o, "spt"), "202");
    write_attribute("path", "m0,0l0,21600,21600,21600,21600,0xe");
    write_start_element(xmlns_v, "stroke");
    write_attribute("joinstyle", "miter");
    write_end_element(xmlns_v, "stroke");
    write_start_element(xmlns_v, "path");
    write_attribute("gradientshapeok", "t");
    write_attribute(xml::qname(xmlns_o, "connecttype"), "rect");
    write_end_element(xmlns_v, "path");
    write_end_element(xmlns_v, "shapetype");

    std::size_t comment_index = 0;

    for (const auto &cell_ref : cells)
    {
        auto comment = ws.cell(cell_ref).comment();
        auto shape_id = 1024 * file_index + 1 + comment_index * 2;

        write_start_element(xmlns_v, "shape");
        write_attribute("id", "_x0000_s" + std::to_string(shape_id));
        write_attribute("type", "#_x0000_t202");

        std::vector<std::pair<std::string, std::string>> style;

        style.push_back({"position", "absolute"});
        style.push_back({"margin-left", std::to_string(comment.left()) + "pt"});
        style.push_back({"margin-top", std::to_string(comment.top()) + "pt"});
        style.push_back({"width", std::to_string(comment.width()) + "pt"});
        style.push_back({"height", std::to_string(comment.height()) + "pt"});
        style.push_back({"z-index", std::to_string(comment_index + 1)});
        style.push_back({"visibility", comment.visible() ? "visible" : "hidden"});

        std::string style_string;

        for (auto part : style)
        {
            style_string.append(part.first);
            style_string.append(":");
            style_string.append(part.second);
            style_string.append(";");
        }

        write_attribute("style", style_string);
        write_attribute("fillcolor", "#fbf6d6");
        write_attribute("strokecolor", "#edeaa1");

        write_start_element(xmlns_v, "fill");
        write_attribute("color2", "#fbfe82");
        write_attribute("angle", -180);
        write_attribute("type", "gradient");
        write_start_element(xmlns_o, "fill");
        write_attribute(xml::qname(xmlns_v, "ext"), "view");
        write_attribute("type", "gradientUnscaled");
        write_end_element(xmlns_o, "fill");
        write_end_element(xmlns_v, "fill");

        write_start_element(xmlns_v, "shadow");
        write_attribute("on", "t");
        write_attribute("obscured", "t");
        write_end_element(xmlns_v, "shadow");

        write_start_element(xmlns_v, "path");
        write_attribute(xml::qname(xmlns_o, "connecttype"), "none");
        write_end_element(xmlns_v, "path");

        write_start_element(xmlns_v, "textbox");
        write_attribute("style", "mso-direction-alt:auto");
        write_start_element("div");
        write_attribute("style", "text-align:left");
        write_characters("");
        write_end_element("div");
        write_end_element(xmlns_v, "textbox");

        write_start_element(xmlns_x, "ClientData");
        write_attribute("ObjectType", "Note");
        write_start_element(xmlns_x, "MoveWithCells");
        write_end_element(xmlns_x, "MoveWithCells");
        write_start_element(xmlns_x, "SizeWithCells");
        write_end_element(xmlns_x, "SizeWithCells");
        write_start_element(xmlns_x, "Anchor");
        write_characters("1, 15, 0, " + std::to_string(2 + comment_index * 4) + ", 2, 54, 4, 14");
        write_end_element(xmlns_x, "Anchor");
        write_start_element(xmlns_x, "AutoFill");
        write_characters("False");
        write_end_element(xmlns_x, "AutoFill");
        write_start_element(xmlns_x, "Row");
        write_characters(cell_ref.row() - 1);
        write_end_element(xmlns_x, "Row");
        write_start_element(xmlns_x, "Column");
        write_characters(cell_ref.column_index() - 1);
        write_end_element(xmlns_x, "Column");
        write_end_element(xmlns_x, "ClientData");

        write_end_element(xmlns_v, "shape");

        ++comment_index;
    }

    write_end_element("xml");
}

// Other Parts

void xlsx_producer::write_custom_property()
{
}

void xlsx_producer::write_unknown_parts()
{
}

void xlsx_producer::write_unknown_relationships()
{
}

void xlsx_producer::write_image(const path &image_path)
{
    end_part();

    vector_istreambuf buffer(source_.d_->images_.at(image_path.string()));
    auto image_streambuf = archive_->open(image_path);
    std::ostream(image_streambuf.get()) << &buffer;
}

std::string xlsx_producer::write_bool(bool boolean) const
{
    return boolean ? "1" : "0";
}

void xlsx_producer::write_relationships(const std::vector<xlnt::relationship> &relationships, const path &part)
{
    path parent = part.parent();

    if (parent.is_absolute())
    {
        parent = path(parent.string().substr(1));
    }

    path rels_path(parent.append("_rels").append(part.filename() + ".rels").string());
    begin_part(rels_path);

    const auto xmlns = xlnt::constants::ns("relationships");

    write_start_element(xmlns, "Relationships");
    write_namespace(xmlns, "");

    for (std::size_t i = 1; i <= relationships.size(); ++i)
    {
        auto rel_iter = std::find_if(relationships.begin(), relationships.end(),
            [&i](const relationship &r) { return r.id() == "rId" + std::to_string(i); });
        auto relationship = *rel_iter;

        write_start_element(xmlns, "Relationship");

        write_attribute("Id", relationship.id());
        write_attribute("Type", relationship.type());
        write_attribute("Target", relationship.target().path().string());

        if (relationship.target_mode() == xlnt::target_mode::external)
        {
            write_attribute("TargetMode", "External");
        }

        write_end_element(xmlns, "Relationship");
    }

    write_end_element(xmlns, "Relationships");
}

void xlsx_producer::write_color(const xlnt::color &color)
{
    if (color.auto_())
    {
        write_attribute("auto", write_bool(true));
        return;
    }

    switch (color.type())
    {
    case xlnt::color_type::theme:
        write_attribute("theme", color.theme().index());
        break;

    case xlnt::color_type::indexed:
        write_attribute("indexed", color.indexed().index());
        break;

    case xlnt::color_type::rgb:
        write_attribute("rgb", color.rgb().hex_string());
        break;
    }
}

void xlsx_producer::write_start_element(const std::string &name)
{
    current_part_serializer_->start_element(name);
}

void xlsx_producer::write_start_element(const std::string &ns, const std::string &name)
{
    current_part_serializer_->start_element(ns, name);
}

void xlsx_producer::write_end_element(const std::string &name)
{
    current_part_serializer_->end_element(name);
}

void xlsx_producer::write_end_element(const std::string &ns, const std::string &name)
{
    current_part_serializer_->end_element(ns, name);
}

void xlsx_producer::write_namespace(const std::string &ns, const std::string &prefix)
{
    current_part_serializer_->namespace_decl(ns, prefix);
}

} // namespace detail
} // namepsace xlnt
