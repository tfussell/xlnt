// Copyright (c) 2014-2021 Thomas Fussell
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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#include <cassert>
#include <cctype>
#include <numeric> // for std::accumulate
#include <sstream>
#include <unordered_map>

#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/comment.hpp>
#include <xlnt/cell/hyperlink.hpp>
#include <xlnt/drawing/spreadsheet_drawing.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/utils/optional.hpp>
#include <xlnt/utils/path.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/selection.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <detail/constants.hpp>
#include <detail/header_footer/header_footer_code.hpp>
#include <detail/implementations/workbook_impl.hpp>
#include <detail/serialization/custom_value_traits.hpp>
#include <detail/serialization/defined_name.hpp>
#include <detail/serialization/serialisation_helpers.hpp>
#include <detail/serialization/vector_streambuf.hpp>
#include <detail/serialization/xlsx_consumer.hpp>
#include <detail/serialization/zstream.hpp>

namespace {
/// string_equal
/// for comparison between std::string and string literals
/// improves on std::string::operator==(char*) by knowing the length ahead of time
template <size_t N>
inline bool string_arr_loop_equal(const std::string &lhs, const char (&rhs)[N])
{
    for (size_t i = 0; i < N - 1; ++i)
    {
        if (lhs[i] != rhs[i])
        {
            return false;
        }
    }
    return true;
}

template <size_t N>
inline bool string_equal(const std::string &lhs, const char (&rhs)[N])
{
    if (lhs.size() != N - 1)
    {
        return false;
    }
    // split function to assist with inlining of the size check
    return string_arr_loop_equal(lhs, rhs);
}

xml::qname &qn(const std::string &namespace_, const std::string &name)
{
    using qname_map = std::unordered_map<std::string, xml::qname>;
    static auto memo = std::unordered_map<std::string, qname_map>();

    auto &ns_memo = memo[namespace_];

    if (ns_memo.find(name) == ns_memo.end())
    {
        return ns_memo.emplace(name, xml::qname(xlnt::constants::ns(namespace_), name)).first->second;
    }

    return ns_memo[name];
}

/// <summary>
/// Returns true if bool_string represents a true xsd:boolean.
/// </summary>
bool is_true(const std::string &bool_string)
{
    if (bool_string == "1" || bool_string == "true")
    {
        return true;
    }

#ifdef THROW_ON_INVALID_XML
    if (bool_string == "0" || bool_string == "false")
    {
        return false;
    }

    throw xlnt::exception("xsd:boolean should be one of: 0, 1, true, or false, found " + bool_string);
#else

    return false;
#endif
}

using style_id_pair = std::pair<xlnt::detail::style_impl, std::size_t>;

/// <summary>
/// Try to find given xfid value in the styles vector and, if succeeded, set's the optional style.
/// </summary>
void set_style_by_xfid(const std::vector<style_id_pair> &styles,
    std::size_t xfid, xlnt::optional<std::string> &style)
{
    for (auto &item : styles)
    {
        if (item.second == xfid)
        {
            style = item.first.name;
        }
    }
}

// <sheetData> element
struct Sheet_Data
{
    std::vector<std::pair<xlnt::row_properties, xlnt::row_t>> parsed_rows;
    std::vector<xlnt::detail::Cell> parsed_cells;
};

xlnt::cell_type type_from_string(const std::string &str)
{
    if (string_equal(str, "s"))
    {
        return xlnt::cell::type::shared_string;
    }
    else if (string_equal(str, "n"))
    {
        return xlnt::cell::type::number;
    }
    else if (string_equal(str, "b"))
    {
        return xlnt::cell::type::boolean;
    }
    else if (string_equal(str, "e"))
    {
        return xlnt::cell::type::error;
    }
    else if (string_equal(str, "inlineStr"))
    {
        return xlnt::cell::type::inline_string;
    }
    else if (string_equal(str, "str"))
    {
        return xlnt::cell::type::formula_string;
    }
    return xlnt::cell::type::shared_string;
}

xlnt::detail::Cell parse_cell(xlnt::row_t row_arg, xml::parser *parser, std::unordered_map<std::string, std::string> &array_formulae, std::unordered_map<int, std::string> &shared_formulae)
{
    xlnt::detail::Cell c;
    for (auto &attr : parser->attribute_map())
    {
        if (string_equal(attr.first.name(), "r"))
        {
            c.ref = xlnt::detail::Cell_Reference(row_arg, attr.second.value);
        }
        else if (string_equal(attr.first.name(), "t"))
        {
            c.type = type_from_string(attr.second.value);
        }
        else if (string_equal(attr.first.name(), "s"))
        {
            c.style_index = static_cast<int>(strtol(attr.second.value.c_str(), nullptr, 10));
        }
        else if (string_equal(attr.first.name(), "ph"))
        {
            c.is_phonetic = is_true(attr.second.value);
        }
        else if (string_equal(attr.first.name(), "cm"))
        {
            c.cell_metatdata_idx = static_cast<int>(strtol(attr.second.value.c_str(), nullptr, 10));
        }
    }
    int level = 1; // nesting level
        // 1 == <c>
        // 2 == <v>/<f>
        // 3 == <is><t>
        // exit loop at </c>
    while (level > 0)
    {
        xml::parser::event_type e = parser->next();
        switch (e)
        {
        case xml::parser::start_element: {
            if (string_equal(parser->name(), "f") && parser->attribute_present("t"))
            {
                // Skip shared formulas with a ref attribute because it indicates that this
                // is the master cell which will be handled in the xml::parser::characters case.
                if (parser->attribute("t") == "shared" && !parser->attribute_present("ref"))
                {
                    auto shared_index = parser->attribute<int>("si");
                    c.formula_string = shared_formulae[shared_index];
                }
            }
            ++level;
            break;
        }
        case xml::parser::end_element: {
            --level;
            break;
        }
        case xml::parser::characters: {
            // only want the characters inside one of the nested tags
            // without this a lot of formatting whitespace can get added
            if (level == 2)
            {
                // <v> -> numeric values
                if (string_equal(parser->name(), "v"))
                {
                    c.value += std::move(parser->value());
                }
                // <f> formula
                else if (string_equal(parser->name(), "f"))
                {
                    c.formula_string += std::move(parser->value());
                    
                    if (parser->attribute_present("t"))
                    {
                        auto formula_ref = parser->attribute("ref");
                        auto formula_type = parser->attribute("t");
                        if (formula_type == "shared")
                        {
                            auto shared_index = parser->attribute<int>("si");
                            shared_formulae[shared_index] = c.formula_string;
                        }
                        else if (formula_type == "array")
                        {
                            array_formulae[formula_ref] = c.formula_string;
                        }
                    }
                }
            }
            else if (level == 3)
            {
                // <is><t> -> inline string
                if (string_equal(parser->name(), "t"))
                {
                    c.value += std::move(parser->value());
                }
            }
            break;
        }
        case xml::parser::start_namespace_decl:
        case xml::parser::end_namespace_decl:
        case xml::parser::start_attribute:
        case xml::parser::end_attribute:
        case xml::parser::eof:
        default: {
            throw xlnt::exception("unexcpected XML parsing event");
        }
        }
        // Prevents unhandled exceptions from being triggered.
        parser->attribute_map();
    }
    return c;
}

// <row> inside <sheetData> element
std::pair<xlnt::row_properties, int> parse_row(xml::parser *parser, xlnt::detail::number_serialiser &converter, std::vector<xlnt::detail::Cell> &parsed_cells, std::unordered_map<std::string, std::string> &array_formulae, std::unordered_map<int, std::string> &shared_formulae)
{
    std::pair<xlnt::row_properties, int> props;
    for (auto &attr : parser->attribute_map())
    {
        if (string_equal(attr.first.name(), "dyDescent"))
        {
            props.first.dy_descent = converter.deserialise(attr.second.value);
        }
        else if (string_equal(attr.first.name(), "spans"))
        {
            props.first.spans = attr.second.value;
        }
        else if (string_equal(attr.first.name(), "ht"))
        {
            props.first.height = converter.deserialise(attr.second.value);
        }
        else if (string_equal(attr.first.name(), "s"))
        {
            props.first.style = strtoul(attr.second.value.c_str(), nullptr, 10);
        }
        else if (string_equal(attr.first.name(), "hidden"))
        {
            props.first.hidden = is_true(attr.second.value);
        }
        else if (string_equal(attr.first.name(), "customFormat"))
        {
            props.first.custom_format = is_true(attr.second.value);
        }
        else if (string_equal(attr.first.name(), "ph"))
        {
            is_true(attr.second.value);
        }
        else if (string_equal(attr.first.name(), "r"))
        {
            props.second = static_cast<int>(strtol(attr.second.value.c_str(), nullptr, 10));
        }
        else if (string_equal(attr.first.name(), "customHeight"))
        {
            props.first.custom_height = is_true(attr.second.value.c_str());
        }
    }

    int level = 1;
    while (level > 0)
    {
        xml::parser::event_type e = parser->next();
        switch (e)
        {
        case xml::parser::start_element: {
            parsed_cells.push_back(parse_cell(static_cast<xlnt::row_t>(props.second), parser, array_formulae, shared_formulae));
            break;
        }
        case xml::parser::end_element: {
            --level;
            break;
        }
        case xml::parser::characters: {
            // ignore whitespace
            break;
        }
        case xml::parser::start_namespace_decl:
        case xml::parser::start_attribute:
        case xml::parser::end_namespace_decl:
        case xml::parser::end_attribute:
        case xml::parser::eof:
        default: {
            throw xlnt::exception("unexcpected XML parsing event");
        }
        }
    }
    return props;
}

// <sheetData> inside <worksheet> element
Sheet_Data parse_sheet_data(xml::parser *parser, xlnt::detail::number_serialiser &converter, std::unordered_map<std::string, std::string> &array_formulae, std::unordered_map<int, std::string> &shared_formulae)
{
    Sheet_Data sheet_data;
    int level = 1; // nesting level
        // 1 == <sheetData>
        // 2 == <row>

    while (level > 0)
    {
        xml::parser::event_type e = parser->next();
        switch (e)
        {
        case xml::parser::start_element: {
            sheet_data.parsed_rows.push_back(parse_row(parser, converter, sheet_data.parsed_cells, array_formulae, shared_formulae));
            break;
        }
        case xml::parser::end_element: {
            --level;
            break;
        }
        case xml::parser::characters: {
            // ignore, whitespace formatting normally
            break;
        }
        case xml::parser::start_namespace_decl:
        case xml::parser::start_attribute:
        case xml::parser::end_namespace_decl:
        case xml::parser::end_attribute:
        case xml::parser::eof:
        default: {
            throw xlnt::exception("unexcpected XML parsing event");
        }
        }
    }
    return sheet_data;
}

} // namespace

/*
class parsing_context
{
public:
    parsing_context(xlnt::detail::zip_file_reader &archive, const std::string &filename)
        : parser_(stream_, filename)
    {
    }

    xml::parser &parser();

private:
    std::istream stream_;
    xml::parser parser_;
};
*/

namespace xlnt {
namespace detail {

xlsx_consumer::xlsx_consumer(workbook &target)
    : target_(target),
      parser_(nullptr)
{
}

xlsx_consumer::~xlsx_consumer()
{
}

void xlsx_consumer::read(std::istream &source)
{
    archive_.reset(new izstream(source));
    populate_workbook(false);
}

void xlsx_consumer::open(std::istream &source)
{
    archive_.reset(new izstream(source));
    populate_workbook(true);
}

cell xlsx_consumer::read_cell()
{
    return cell(streaming_cell_.get());
}

void xlsx_consumer::read_worksheet(const std::string &rel_id)
{
    read_worksheet_begin(rel_id);

    if (!streaming_)
    {
        read_worksheet_sheetdata();
        read_worksheet_end(rel_id);
    }
}

void read_defined_names(worksheet ws, std::vector<defined_name> defined_names)
{
    for (auto &name : defined_names)
    {
        if (name.sheet_id != ws.id() - 1)
        {
            continue;
        }
        
        if (name.name == "_xlnm.Print_Titles")
        {
            // Basic print titles parser
            // A print title definition looks like "'Sheet3'!$B:$E,'Sheet3'!$2:$4"
            // There are three cases: columns only, rows only, and both (separated by a comma).
            // For this reason, we loop up to two times parsing each component.
            // Titles may be quoted (with single quotes) or unquoted. We ignore them for now anyways.
            // References are always absolute.
            // Move this into a separate function if it needs to be used in other places.
            auto i = std::size_t(0);
            for (auto count = 0; count < 2; count++)
            {
                // Split into components based on "!", ":", and "," characters
                auto j = i;
                i = name.value.find("!", j);
                auto title = name.value.substr(j, i - j);
                j = i + 2; // skip "!$"
                i = name.value.find(":", j);
                auto from = name.value.substr(j, i - j);
                j = i + 2; // skip ":$"
                i = name.value.find(",", j);
                auto to = name.value.substr(j, i - j);

                // Apply to the worksheet
                if (isalpha(from.front())) // alpha=>columns
                {
                    ws.print_title_cols(from, to);
                }
                else // numeric=>rows
                {
                    ws.print_title_rows(std::stoul(from), std::stoul(to));
                }
                
                // Check for end condition
                if (i == std::string::npos)
                {
                    break;
                }

                i++; // skip "," for next iteration
            }
        }
        else if (name.name == "_xlnm._FilterDatabase")
        {
            auto i = name.value.find("!");
            ws.auto_filter(name.value.substr(i + 1));
        }
        else if (name.name == "_xlnm.Print_Area")
        {
            auto i = name.value.find("!");
            ws.print_area(name.value.substr(i + 1));
        }
    }
}

std::string xlsx_consumer::read_worksheet_begin(const std::string &rel_id)
{
    if (streaming_ && streaming_cell_ == nullptr)
    {
        streaming_cell_.reset(new detail::cell_impl());
    }
    
    array_formulae_.clear();
    shared_formulae_.clear();

    auto title = std::find_if(target_.d_->sheet_title_rel_id_map_.begin(),
        target_.d_->sheet_title_rel_id_map_.end(),
        [&](const std::pair<std::string, std::string> &p) {
            return p.second == rel_id;
        })->first;

    auto ws = worksheet(current_worksheet_);

    expect_start_element(qn("spreadsheetml", "worksheet"), xml::content::complex); // CT_Worksheet
    skip_attributes({qn("mc", "Ignorable")});
    
    read_defined_names(ws, defined_names_);

    while (in_element(qn("spreadsheetml", "worksheet")))
    {
        auto current_worksheet_element = expect_start_element(xml::content::complex);

        if (current_worksheet_element == qn("spreadsheetml", "sheetPr")) // CT_SheetPr 0-1
        {
            sheet_pr props;
            if (parser().attribute_present("syncHorizontal"))
            { // optional, boolean, false
                props.sync_horizontal.set(parser().attribute<bool>("syncHorizontal"));
            }
            if (parser().attribute_present("syncVertical"))
            { // optional, boolean, false
                props.sync_vertical.set(parser().attribute<bool>("syncVertical"));
            }
            if (parser().attribute_present("syncRef"))
            { // optional, ST_Ref, false
                props.sync_ref.set(cell_reference(parser().attribute("syncRef")));
            }
            if (parser().attribute_present("transitionEvaluation"))
            { // optional, boolean, false
                props.transition_evaluation.set(parser().attribute<bool>("transitionEvaluation"));
            }
            if (parser().attribute_present("transitionEntry"))
            { // optional, boolean, false
                props.transition_entry.set(parser().attribute<bool>("transitionEntry"));
            }
            if (parser().attribute_present("published"))
            { // optional, boolean, true
                props.published.set(parser().attribute<bool>("published"));
            }
            if (parser().attribute_present("codeName"))
            { // optional, string
                props.code_name.set(parser().attribute<std::string>("codeName"));
            }
            if (parser().attribute_present("filterMode"))
            { // optional, boolean, false
                props.filter_mode.set(parser().attribute<bool>("filterMode"));
            }
            if (parser().attribute_present("enableFormatConditionsCalculation"))
            { // optional, boolean, true
                props.enable_format_condition_calculation.set(parser().attribute<bool>("enableFormatConditionsCalculation"));
            }
            ws.d_->sheet_properties_.set(props);
            while (in_element(current_worksheet_element))
            {
                auto sheet_pr_child_element = expect_start_element(xml::content::simple);

                if (sheet_pr_child_element == qn("spreadsheetml", "tabColor")) // CT_Color 0-1
                {
                    read_color();
                }
                else if (sheet_pr_child_element == qn("spreadsheetml", "outlinePr")) // CT_OutlinePr 0-1
                {
                    skip_attribute("applyStyles"); // optional, boolean, false
                    skip_attribute("summaryBelow"); // optional, boolean, true
                    skip_attribute("summaryRight"); // optional, boolean, true
                    skip_attribute("showOutlineSymbols"); // optional, boolean, true
                }
                else if (sheet_pr_child_element == qn("spreadsheetml", "pageSetUpPr")) // CT_PageSetUpPr 0-1
                {
                    skip_attribute("autoPageBreaks"); // optional, boolean, true
                    skip_attribute("fitToPage"); // optional, boolean, false
                }
                else
                {
                    unexpected_element(sheet_pr_child_element);
                }

                expect_end_element(sheet_pr_child_element);
            }
        }
        else if (current_worksheet_element == qn("spreadsheetml", "dimension")) // CT_SheetDimension 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "sheetViews")) // CT_SheetViews 0-1
        {
            while (in_element(current_worksheet_element))
            {
                expect_start_element(qn("spreadsheetml", "sheetView"), xml::content::complex); // CT_SheetView 1+

                sheet_view new_view;
                new_view.id(parser().attribute<std::size_t>("workbookViewId"));

                if (parser().attribute_present("showGridLines")) // default="true"
                {
                    new_view.show_grid_lines(is_true(parser().attribute("showGridLines")));
                }
                if (parser().attribute_present("topLeftCell"))
                {
                    new_view.top_left_cell(cell_reference(parser().attribute("topLeftCell")));
                }

                if (parser().attribute_present("defaultGridColor")) // default="true"
                {
                    new_view.default_grid_color(is_true(parser().attribute("defaultGridColor")));
                }

                if (parser().attribute_present("view")
                    && parser().attribute("view") != "normal")
                {
                    new_view.type(parser().attribute("view") == "pageBreakPreview"
                            ? sheet_view_type::page_break_preview
                            : sheet_view_type::page_layout);
                }

                if (parser().attribute_present("tabSelected")
                    && is_true(parser().attribute("tabSelected")))
                {
                    target_.d_->view_.get().active_tab = ws.id() - 1;
                }

                skip_attributes({"windowProtection", "showFormulas", "showRowColHeaders", "showZeros", "rightToLeft", "showRuler", "showOutlineSymbols", "showWhiteSpace",
                    "view", "topLeftCell", "colorId", "zoomScale", "zoomScaleNormal", "zoomScaleSheetLayoutView",
                    "zoomScalePageLayoutView"});

                while (in_element(qn("spreadsheetml", "sheetView")))
                {
                    auto sheet_view_child_element = expect_start_element(xml::content::simple);

                    if (sheet_view_child_element == qn("spreadsheetml", "pane")) // CT_Pane 0-1
                    {
                        pane new_pane;

                        if (parser().attribute_present("topLeftCell"))
                        {
                            new_pane.top_left_cell = cell_reference(parser().attribute("topLeftCell"));
                        }

                        if (parser().attribute_present("xSplit"))
                        {
                            new_pane.x_split = parser().attribute<column_t::index_t>("xSplit");
                        }

                        if (parser().attribute_present("ySplit"))
                        {
                            new_pane.y_split = parser().attribute<row_t>("ySplit");
                        }

                        if (parser().attribute_present("activePane"))
                        {
                            new_pane.active_pane = parser().attribute<pane_corner>("activePane");
                        }

                        if (parser().attribute_present("state"))
                        {
                            new_pane.state = parser().attribute<pane_state>("state");
                        }

                        new_view.pane(new_pane);
                    }
                    else if (sheet_view_child_element == qn("spreadsheetml", "selection")) // CT_Selection 0-4
                    {
                        selection current_selection;

                        if (parser().attribute_present("activeCell"))
                        {
                            current_selection.active_cell(parser().attribute("activeCell"));
                        }

                        if (parser().attribute_present("sqref"))
                        {
                            const auto sqref = range_reference(parser().attribute("sqref"));
                            current_selection.sqref(sqref);
                        }

                        if (parser().attribute_present("pane"))
                        {
                            current_selection.pane(parser().attribute<pane_corner>("pane"));
                        }

                        new_view.add_selection(current_selection);

                        skip_remaining_content(sheet_view_child_element);
                    }
                    else if (sheet_view_child_element == qn("spreadsheetml", "pivotSelection")) // CT_PivotSelection 0-4
                    {
                        skip_remaining_content(sheet_view_child_element);
                    }
                    else if (sheet_view_child_element == qn("spreadsheetml", "extLst")) // CT_ExtensionList 0-1
                    {
                        skip_remaining_content(sheet_view_child_element);
                    }
                    else
                    {
                        unexpected_element(sheet_view_child_element);
                    }

                    expect_end_element(sheet_view_child_element);
                }

                expect_end_element(qn("spreadsheetml", "sheetView"));

                ws.d_->views_.push_back(new_view);
            }
        }
        else if (current_worksheet_element == qn("spreadsheetml", "sheetFormatPr")) // CT_SheetFormatPr 0-1
        {
            if (parser().attribute_present("baseColWidth"))
            {
                ws.d_->format_properties_.base_col_width =
                    converter_.deserialise(parser().attribute("baseColWidth"));
            }
            if (parser().attribute_present("defaultColWidth"))
            {
                ws.d_->format_properties_.default_column_width =
                    converter_.deserialise(parser().attribute("defaultColWidth"));
            }
            if (parser().attribute_present("defaultRowHeight"))
            {
                ws.d_->format_properties_.default_row_height =
                    converter_.deserialise(parser().attribute("defaultRowHeight"));
            }

            if (parser().attribute_present(qn("x14ac", "dyDescent")))
            {
                ws.d_->format_properties_.dy_descent =
                    converter_.deserialise(parser().attribute(qn("x14ac", "dyDescent")));
            }

            skip_attributes();
        }
        else if (current_worksheet_element == qn("spreadsheetml", "cols")) // CT_Cols 0+
        {
            while (in_element(qn("spreadsheetml", "cols")))
            {
                expect_start_element(qn("spreadsheetml", "col"), xml::content::simple);

                skip_attributes(std::vector<std::string>{"collapsed", "outlineLevel"});

                auto min = static_cast<column_t::index_t>(std::stoull(parser().attribute("min")));
                auto max = static_cast<column_t::index_t>(std::stoull(parser().attribute("max")));

                // avoid uninitialised warnings in GCC by using a lambda to make the conditional initialisation
                optional<double> width = [this](xml::parser &p) -> xlnt::optional<double> {
                    if (p.attribute_present("width"))
                    {
                        return (converter_.deserialise(p.attribute("width")) * 7 - 5) / 7;
                    }
                    return xlnt::optional<double>();
                }(parser());
                // avoid uninitialised warnings in GCC by using a lambda to make the conditional initialisation
                optional<std::size_t> column_style = [](xml::parser &p) -> xlnt::optional<std::size_t> {
                    if (p.attribute_present("style"))
                    {
                        return p.attribute<std::size_t>("style");
                    }
                    return xlnt::optional<std::size_t>();
                }(parser());

                auto custom = parser().attribute_present("customWidth")
                    ? is_true(parser().attribute("customWidth"))
                    : false;
                auto hidden = parser().attribute_present("hidden")
                    ? is_true(parser().attribute("hidden"))
                    : false;
                auto best_fit = parser().attribute_present("bestFit")
                    ? is_true(parser().attribute("bestFit"))
                    : false;

                expect_end_element(qn("spreadsheetml", "col"));

                for (auto column = min; column <= max; column++)
                {
                    column_properties props;

                    if (width.is_set())
                    {
                        props.width = width.get();
                    }

                    if (column_style.is_set())
                    {
                        props.style = column_style.get();
                    }

                    props.hidden = hidden;
                    props.custom_width = custom;
                    props.best_fit = best_fit;
                    ws.add_column_properties(column, props);
                }
            }
        }
        else if (current_worksheet_element == qn("spreadsheetml", "sheetData")) // CT_SheetData 1
        {
            return title;
        }

        expect_end_element(current_worksheet_element);
    }

    return title;
}

void xlsx_consumer::read_worksheet_sheetdata()
{
    if (stack_.back() != qn("spreadsheetml", "sheetData"))
    {
        return;
    }

    auto ws_data = parse_sheet_data(parser_, converter_, array_formulae_, shared_formulae_);
    // NOTE: parse->construct are seperated here and could easily be threaded
    // with a SPSC queue for what is likely to be an easy performance win
    for (auto &row : ws_data.parsed_rows)
    {
        current_worksheet_->row_properties_.emplace(row.second, std::move(row.first));
    }
    auto impl = detail::cell_impl();
    for (Cell &cell : ws_data.parsed_cells)
    {
        impl.parent_ = current_worksheet_;
        impl.column_ = cell.ref.column;
        impl.row_ = cell.ref.row;
        detail::cell_impl *ws_cell_impl = &current_worksheet_->cell_map_.emplace(cell_reference(impl.column_, impl.row_), std::move(impl)).first->second;
        if (cell.style_index != -1)
        {
            ws_cell_impl->format_ = target_.format(static_cast<size_t>(cell.style_index)).d_;
        }
        if (cell.cell_metatdata_idx != -1)
        {
        }
        ws_cell_impl->phonetics_visible_ = cell.is_phonetic;
        if (!cell.formula_string.empty())
        {
            ws_cell_impl->formula_ = cell.formula_string[0] == '=' ? cell.formula_string.substr(1) : std::move(cell.formula_string);
        }
        if (!cell.value.empty())
        {
            ws_cell_impl->type_ = cell.type;
            switch (cell.type)
            {
            case cell::type::boolean: {
                ws_cell_impl->value_numeric_ = is_true(cell.value) ? 1.0 : 0.0;
                break;
            }
            case cell::type::empty:
            case cell::type::number:
            case cell::type::date: {
                ws_cell_impl->value_numeric_ = converter_.deserialise(cell.value);
                break;
            }
            case cell::type::shared_string: {
                ws_cell_impl->value_numeric_ = static_cast<double>(strtol(cell.value.c_str(), nullptr, 10));
                break;
            }
            case cell::type::inline_string: {
                ws_cell_impl->value_text_ = std::move(cell.value);
                break;
            }
            case cell::type::formula_string: {
                ws_cell_impl->value_text_ = std::move(cell.value);
                break;
            }
            case cell::type::error: {
                ws_cell_impl->value_text_.plain_text(cell.value, false);
                break;
            }
            }
        }
    }
    stack_.pop_back();
    
    
}

worksheet xlsx_consumer::read_worksheet_end(const std::string &rel_id)
{
    auto &manifest = target_.manifest();

    const auto workbook_rel = manifest.relationship(path("/"), relationship_type::office_document);
    const auto sheet_rel = manifest.relationship(workbook_rel.target().path(), rel_id);
    path sheet_path(sheet_rel.source().path().parent().append(sheet_rel.target().path()));
    auto hyperlinks = manifest.relationships(sheet_path, xlnt::relationship_type::hyperlink);

    auto ws = worksheet(current_worksheet_);

    while (in_element(qn("spreadsheetml", "worksheet")))
    {
        auto current_worksheet_element = expect_start_element(xml::content::complex);

        if (current_worksheet_element == qn("spreadsheetml", "sheetCalcPr")) // CT_SheetCalcPr 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "sheetProtection")) // CT_SheetProtection 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "protectedRanges")) // CT_ProtectedRanges 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "scenarios")) // CT_Scenarios 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "autoFilter")) // CT_AutoFilter 0-1
        {
            ws.auto_filter(xlnt::range_reference(parser().attribute("ref")));
            // auto filter complex
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "sortState")) // CT_SortState 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "dataConsolidate")) // CT_DataConsolidate 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "customSheetViews")) // CT_CustomSheetViews 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "mergeCells")) // CT_MergeCells 0-1
        {
            parser().attribute_map();

            while (in_element(qn("spreadsheetml", "mergeCells")))
            {
                expect_start_element(qn("spreadsheetml", "mergeCell"), xml::content::simple);
                ws.merge_cells(range_reference(parser().attribute("ref")));
                expect_end_element(qn("spreadsheetml", "mergeCell"));
            }
        }
        else if (current_worksheet_element == qn("spreadsheetml", "phoneticPr")) // CT_PhoneticPr 0-1
        {
            phonetic_pr phonetic_properties(parser().attribute<std::uint32_t>("fontId"));
            if (parser().attribute_present("type"))
            {
                phonetic_properties.type(phonetic_pr::type_from_string(parser().attribute("type")));
            }
            if (parser().attribute_present("alignment"))
            {
                phonetic_properties.alignment(phonetic_pr::alignment_from_string(parser().attribute("alignment")));
            }
            current_worksheet_->phonetic_properties_.set(phonetic_properties);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "conditionalFormatting")) // CT_ConditionalFormatting 0+
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "dataValidations")) // CT_DataValidations 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "hyperlinks")) // CT_Hyperlinks 0-1
        {
            while (in_element(current_worksheet_element))
            {
                // CT_Hyperlink
                expect_start_element(qn("spreadsheetml", "hyperlink"), xml::content::simple);

                auto cell = ws.cell(parser().attribute("ref"));

                if (parser().attribute_present(qn("r", "id")))
                {
                    auto hyperlink_rel_id = parser().attribute(qn("r", "id"));
                    auto hyperlink_rel = std::find_if(hyperlinks.begin(), hyperlinks.end(),
                        [&](const relationship &r) { return r.id() == hyperlink_rel_id; });

                    if (hyperlink_rel != hyperlinks.end())
                    {
                        auto url = hyperlink_rel->target().path().string();

                        if (cell.has_value())
                        {
                            cell.hyperlink(url, cell.value<std::string>());
                        }
                        else
                        {
                            cell.hyperlink(url);
                        }
                    }
                }
                else if (parser().attribute_present("location"))
                {
                    auto hyperlink = hyperlink_impl();

                    auto location = parser().attribute("location");
                    hyperlink.relationship = relationship("", relationship_type::hyperlink,
                        uri(""), uri(location), target_mode::internal);

                    if (parser().attribute_present("display"))
                    {
                        hyperlink.display = parser().attribute("display");
                    }

                    if (parser().attribute_present("tooltip"))
                    {
                        hyperlink.tooltip = parser().attribute("tooltip");
                    }

                    cell.d_->hyperlink_ = hyperlink;
                }

                expect_end_element(qn("spreadsheetml", "hyperlink"));
            }
        }
        else if (current_worksheet_element == qn("spreadsheetml", "printOptions")) // CT_PrintOptions 0-1
        {
            print_options opts;
            if (parser().attribute_present("gridLines"))
            {
                opts.print_grid_lines.set(parser().attribute<bool>("gridLines"));
            }
            if (parser().attribute_present("gridLinesSet"))
            {
                opts.grid_lines_set.set(parser().attribute<bool>("gridLinesSet"));
            }
            if (parser().attribute_present("headings"))
            {
                opts.print_headings.set(parser().attribute<bool>("headings"));
            }
            if (parser().attribute_present("horizontalCentered"))
            {
                opts.horizontal_centered.set(parser().attribute<bool>("horizontalCentered"));
            }
            if (parser().attribute_present("verticalCentered"))
            {
                opts.vertical_centered.set(parser().attribute<bool>("verticalCentered"));
            }
            ws.d_->print_options_.set(opts);
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "pageMargins")) // CT_PageMargins 0-1
        {
            page_margins margins;

            margins.top(converter_.deserialise(parser().attribute("top")));
            margins.bottom(converter_.deserialise(parser().attribute("bottom")));
            margins.left(converter_.deserialise(parser().attribute("left")));
            margins.right(converter_.deserialise(parser().attribute("right")));
            margins.header(converter_.deserialise(parser().attribute("header")));
            margins.footer(converter_.deserialise(parser().attribute("footer")));

            ws.page_margins(margins);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "pageSetup")) // CT_PageSetup 0-1
        {
            page_setup setup;
            if (parser().attribute_present("orientation"))
            {
                setup.orientation_.set(parser().attribute<orientation>("orientation"));
            }
            if (parser().attribute_present("horizontalDpi"))
            {
                setup.horizontal_dpi_.set(parser().attribute<std::size_t>("horizontalDpi"));
            }
            if (parser().attribute_present("verticalDpi"))
            {
                setup.vertical_dpi_.set(parser().attribute<std::size_t>("verticalDpi"));
            }
            if (parser().attribute_present("paperSize"))
            {
                setup.paper_size(static_cast<xlnt::paper_size>(parser().attribute<std::size_t>("paperSize")));
            }
            if (parser().attribute_present("scale"))
            {
                setup.scale(parser().attribute<double>("scale"));
            }
            if (parser().attribute_present(qn("r", "id")))
            {
                setup.rel_id(parser().attribute(qn("r", "id")));
            }
            ws.page_setup(setup);
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "headerFooter")) // CT_HeaderFooter 0-1
        {
            header_footer hf;

            hf.align_with_margins(!parser().attribute_present("alignWithMargins")
                || is_true(parser().attribute("alignWithMargins")));
            hf.scale_with_doc(!parser().attribute_present("alignWithMargins")
                || is_true(parser().attribute("alignWithMargins")));
            auto different_odd_even = parser().attribute_present("differentOddEven")
                && is_true(parser().attribute("differentOddEven"));
            auto different_first = parser().attribute_present("differentFirst")
                && is_true(parser().attribute("differentFirst"));

            optional<std::array<optional<rich_text>, 3>> odd_header;
            optional<std::array<optional<rich_text>, 3>> odd_footer;
            optional<std::array<optional<rich_text>, 3>> even_header;
            optional<std::array<optional<rich_text>, 3>> even_footer;
            optional<std::array<optional<rich_text>, 3>> first_header;
            optional<std::array<optional<rich_text>, 3>> first_footer;

            using xlnt::detail::decode_header_footer;

            while (in_element(current_worksheet_element))
            {
                auto current_hf_element = expect_start_element(xml::content::simple);

                if (current_hf_element == qn("spreadsheetml", "oddHeader"))
                {
                    odd_header = decode_header_footer(read_text(), converter_);
                }
                else if (current_hf_element == qn("spreadsheetml", "oddFooter"))
                {
                    odd_footer = decode_header_footer(read_text(), converter_);
                }
                else if (current_hf_element == qn("spreadsheetml", "evenHeader"))
                {
                    even_header = decode_header_footer(read_text(), converter_);
                }
                else if (current_hf_element == qn("spreadsheetml", "evenFooter"))
                {
                    even_footer = decode_header_footer(read_text(), converter_);
                }
                else if (current_hf_element == qn("spreadsheetml", "firstHeader"))
                {
                    first_header = decode_header_footer(read_text(), converter_);
                }
                else if (current_hf_element == qn("spreadsheetml", "firstFooter"))
                {
                    first_footer = decode_header_footer(read_text(), converter_);
                }
                else
                {
                    unexpected_element(current_hf_element);
                }

                expect_end_element(current_hf_element);
            }

            for (std::size_t i = 0; i < 3; ++i)
            {
                auto loc = i == 0 ? header_footer::location::left
                                  : i == 1 ? header_footer::location::center : header_footer::location::right;

                if (different_odd_even)
                {
                    if (odd_header.is_set()
                        && odd_header.get().at(i).is_set()
                        && even_header.is_set()
                        && even_header.get().at(i).is_set())
                    {
                        hf.odd_even_header(loc, odd_header.get().at(i).get(), even_header.get().at(i).get());
                    }

                    if (odd_footer.is_set()
                        && odd_footer.get().at(i).is_set()
                        && even_footer.is_set()
                        && even_footer.get().at(i).is_set())
                    {
                        hf.odd_even_footer(loc, odd_footer.get().at(i).get(), even_footer.get().at(i).get());
                    }
                }
                else
                {
                    if (odd_header.is_set() && odd_header.get().at(i).is_set())
                    {
                        hf.header(loc, odd_header.get().at(i).get());
                    }

                    if (odd_footer.is_set() && odd_footer.get().at(i).is_set())
                    {
                        hf.footer(loc, odd_footer.get().at(i).get());
                    }
                }

                if (different_first)
                {
                }
            }

            ws.header_footer(hf);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "rowBreaks")) // CT_PageBreak 0-1
        {
            auto count = parser().attribute_present("count") ? parser().attribute<std::size_t>("count") : 0;
            auto manual_break_count = parser().attribute_present("manualBreakCount")
                ? parser().attribute<std::size_t>("manualBreakCount")
                : 0;

            while (in_element(qn("spreadsheetml", "rowBreaks")))
            {
                expect_start_element(qn("spreadsheetml", "brk"), xml::content::simple);

                if (parser().attribute_present("id"))
                {
                    ws.page_break_at_row(parser().attribute<row_t>("id"));
                    --count;
                }

                if (parser().attribute_present("man") && is_true(parser().attribute("man")))
                {
                    --manual_break_count;
                }

                skip_attributes({"min", "max", "pt"});
                expect_end_element(qn("spreadsheetml", "brk"));
            }
        }
        else if (current_worksheet_element == qn("spreadsheetml", "colBreaks")) // CT_PageBreak 0-1
        {
            auto count = parser().attribute_present("count") ? parser().attribute<std::size_t>("count") : 0;
            auto manual_break_count = parser().attribute_present("manualBreakCount")
                ? parser().attribute<std::size_t>("manualBreakCount")
                : 0;

            while (in_element(qn("spreadsheetml", "colBreaks")))
            {
                expect_start_element(qn("spreadsheetml", "brk"), xml::content::simple);

                if (parser().attribute_present("id"))
                {
                    ws.page_break_at_column(parser().attribute<column_t::index_t>("id"));
                    --count;
                }

                if (parser().attribute_present("man") && is_true(parser().attribute("man")))
                {
                    --manual_break_count;
                }

                skip_attributes({"min", "max", "pt"});
                expect_end_element(qn("spreadsheetml", "brk"));
            }
        }
        else if (current_worksheet_element == qn("spreadsheetml", "customProperties")) // CT_CustomProperties 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "cellWatches")) // CT_CellWatches 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "ignoredErrors")) // CT_IgnoredErrors 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "smartTags")) // CT_SmartTags 0-1
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "drawing")) // CT_Drawing 0-1
        {
            if (parser().attribute_present(qn("r", "id")))
            {
                auto drawing_rel_id = parser().attribute(qn("r", "id"));
                ws.d_->drawing_rel_id_ = drawing_rel_id;
            }
        }
        else if (current_worksheet_element == qn("spreadsheetml", "legacyDrawing"))
        {
            skip_remaining_content(current_worksheet_element);
        }
        else if (current_worksheet_element == qn("spreadsheetml", "extLst"))
        {
            ext_list extensions(parser(), current_worksheet_element.namespace_());
            ws.d_->extension_list_.set(extensions);
        }
        else
        {
            unexpected_element(current_worksheet_element);
        }

        expect_end_element(current_worksheet_element);
    }

    expect_end_element(qn("spreadsheetml", "worksheet"));

    if (manifest.has_relationship(sheet_path, xlnt::relationship_type::comments))
    {
        auto comments_part = manifest.canonicalize({workbook_rel, sheet_rel,
            manifest.relationship(sheet_path, xlnt::relationship_type::comments)});

        auto receive = xml::parser::receive_default;
        auto comments_part_streambuf = archive_->open(comments_part);
        std::istream comments_part_stream(comments_part_streambuf.get());
        xml::parser parser(comments_part_stream, comments_part.string(), receive);
        parser_ = &parser;

        read_comments(ws);

        if (manifest.has_relationship(sheet_path, xlnt::relationship_type::vml_drawing))
        {
            auto vml_drawings_part = manifest.canonicalize({workbook_rel, sheet_rel,
                manifest.relationship(sheet_path, xlnt::relationship_type::vml_drawing)});

            auto vml_drawings_part_streambuf = archive_->open(comments_part);
            std::istream vml_drawings_part_stream(comments_part_streambuf.get());
            xml::parser vml_parser(vml_drawings_part_stream, vml_drawings_part.string(), receive);
            parser_ = &vml_parser;

            read_vml_drawings(ws);
        }
    }

    if (manifest.has_relationship(sheet_path, xlnt::relationship_type::drawings))
    {
        auto drawings_part = manifest.canonicalize({workbook_rel, sheet_rel,
            manifest.relationship(sheet_path, xlnt::relationship_type::drawings)});

        auto receive = xml::parser::receive_default;
        auto drawings_part_streambuf = archive_->open(drawings_part);
        std::istream drawings_part_stream(drawings_part_streambuf.get());
        xml::parser parser(drawings_part_stream, drawings_part.string(), receive);
        parser_ = &parser;

        read_drawings(ws, drawings_part);
    }

    if (manifest.has_relationship(sheet_path, xlnt::relationship_type::printer_settings))
    {
        read_part({workbook_rel, sheet_rel,
            manifest.relationship(sheet_path,
                relationship_type::printer_settings)});
    }
    
    for (auto array_formula : array_formulae_)
    {
        for (auto row : ws.range(array_formula.first))
        {
            for (auto cell : row)
            {
                cell.formula(array_formula.second);
            }
        }
    }

    return ws;
}

xml::parser &xlsx_consumer::parser()
{
    return *parser_;
}

bool xlsx_consumer::has_cell()
{
    auto ws = worksheet(current_worksheet_);

    while (streaming_cell_ // we're not at the end of the file
           && !in_element(qn("spreadsheetml", "row"))) // we're at the end of a row, or between rows
    {
        if (parser().peek() == xml::parser::event_type::end_element
            && stack_.back() == qn("spreadsheetml", "row"))
        {
            // We're at the end of a row.
            expect_end_element(qn("spreadsheetml", "row"));
            // ... and keep parsing.
        }

        if (parser().peek() == xml::parser::event_type::end_element
            && stack_.back() == qn("spreadsheetml", "sheetData"))
        {
            // End of sheet. Mark it by setting streaming_cell_ to nullptr, so we never get here again.
            expect_end_element(qn("spreadsheetml", "sheetData"));
            streaming_cell_.reset(nullptr);
            break;
        }

        expect_start_element(qn("spreadsheetml", "row"), xml::content::complex); // CT_Row
        auto row_index = static_cast<row_t>(std::stoul(parser().attribute("r")));
        auto &row_properties = ws.row_properties(row_index);

        if (parser().attribute_present("ht"))
        {
            row_properties.height = converter_.deserialise(parser().attribute("ht"));
        }

        if (parser().attribute_present("customHeight"))
        {
            row_properties.custom_height = is_true(parser().attribute("customHeight"));
        }

        if (parser().attribute_present("hidden") && is_true(parser().attribute("hidden")))
        {
            row_properties.hidden = true;
        }

        if (parser().attribute_present(qn("x14ac", "dyDescent")))
        {
            row_properties.dy_descent = converter_.deserialise(parser().attribute(qn("x14ac", "dyDescent")));
        }

        if (parser().attribute_present("spans"))
        {
            row_properties.spans = parser().attribute("spans");
        }

        skip_attributes({"customFormat", "s", "customFont",
            "outlineLevel", "collapsed", "thickTop", "thickBot",
            "ph"});
    }

    if (!streaming_cell_)
    {
        // We're at the end of the worksheet
        return false;
    }

    expect_start_element(qn("spreadsheetml", "c"), xml::content::complex);

    assert(streaming_);
    streaming_cell_.reset(new detail::cell_impl()); // Clean cell state - otherwise it might contain information from the previously streamed cell.
    auto cell = xlnt::cell(streaming_cell_.get());
    auto reference = cell_reference(parser().attribute("r"));
    cell.d_->parent_ = current_worksheet_;
    cell.d_->column_ = reference.column_index();
    cell.d_->row_ = reference.row();

    if (parser().attribute_present("ph"))
    {
        cell.d_->phonetics_visible_ = parser().attribute<bool>("ph");
    }

    auto has_type = parser().attribute_present("t");
    auto type = has_type ? parser().attribute("t") : "n";

    if (parser().attribute_present("s"))
    {
        cell.format(target_.format(static_cast<std::size_t>(std::stoull(parser().attribute("s")))));
    }

    auto has_value = false;
    auto value_string = std::string();
    auto formula_string = std::string();

    while (in_element(qn("spreadsheetml", "c")))
    {
        auto current_element = expect_start_element(xml::content::mixed);

        if (current_element == qn("spreadsheetml", "v")) // s:ST_Xstring
        {
            has_value = true;
            value_string = read_text();
        }
        else if (current_element == qn("spreadsheetml", "f")) // CT_CellFormula
        {
            auto has_shared_formula = false;
            auto has_array_formula = false;
            auto is_master_cell = false;
            auto shared_formula_index = 0;
            auto formula_range = range_reference();

            if (parser().attribute_present("t"))
            {
                auto formula_type = parser().attribute("t");
                if (formula_type == "shared")
                {
                    has_shared_formula = true;
                    shared_formula_index = parser().attribute<int>("si");
                    if (parser().attribute_present("ref"))
                    {
                        is_master_cell = true;
                    }
                }
                else if (formula_type == "array")
                {
                    has_array_formula = true;
                    formula_range = range_reference(parser().attribute("ref"));
                    is_master_cell = true;
                }
            }

            skip_attributes({"aca", "dt2D", "dtr", "del1", "del2", "r1",
                "r2", "ca", "bx"});

            formula_string = read_text();
            
            if (is_master_cell)
            {
                if (has_shared_formula)
                {
                    shared_formulae_[shared_formula_index] = formula_string;
                }
                else if (has_array_formula)
                {
                    array_formulae_[formula_range.to_string()] = formula_string;
                }
            }
            else if (has_shared_formula)
            {
                auto shared_formula = shared_formulae_.find(shared_formula_index);
                if (shared_formula != shared_formulae_.end())
                {
                    formula_string = shared_formula->second;
                }
            }
        }
        else if (current_element == qn("spreadsheetml", "is")) // CT_Rst
        {
            expect_start_element(qn("spreadsheetml", "t"), xml::content::simple);
            has_value = true;
            value_string = read_text();
            expect_end_element(qn("spreadsheetml", "t"));
        }
        else
        {
            unexpected_element(current_element);
        }

        expect_end_element(current_element);
    }

    expect_end_element(qn("spreadsheetml", "c"));

    if (!formula_string.empty())
    {
        cell.formula(formula_string);
    }

    if (has_value)
    {
        if (type == "str")
        {
            cell.d_->value_text_ = value_string;
            cell.data_type(cell::type::formula_string);
        }
        else if (type == "inlineStr")
        {
            cell.d_->value_text_ = value_string;
            cell.data_type(cell::type::inline_string);
        }
        else if (type == "s")
        {
            cell.d_->value_numeric_ = converter_.deserialise(value_string);
            cell.data_type(cell::type::shared_string);
        }
        else if (type == "b") // boolean
        {
            cell.value(is_true(value_string));
        }
        else if (type == "n") // numeric
        {
            cell.value(converter_.deserialise(value_string));
        }
        else if (!value_string.empty() && value_string[0] == '#')
        {
            cell.error(value_string);
        }
    }

    return true;
}

std::vector<relationship> xlsx_consumer::read_relationships(const path &part)
{
    const auto part_rels_path = part.parent().append("_rels").append(part.filename() + ".rels").relative_to(path("/"));

    std::vector<xlnt::relationship> relationships;
    if (!archive_->has_file(part_rels_path)) return relationships;

    auto rels_streambuf = archive_->open(part_rels_path);
    std::istream rels_stream(rels_streambuf.get());
    xml::parser parser(rels_stream, part_rels_path.string());
    parser_ = &parser;

    expect_start_element(qn("relationships", "Relationships"), xml::content::complex);

    while (in_element(qn("relationships", "Relationships")))
    {
        expect_start_element(qn("relationships", "Relationship"), xml::content::simple);

        const auto target_mode = parser.attribute_present("TargetMode")
            ? parser.attribute<xlnt::target_mode>("TargetMode")
            : xlnt::target_mode::internal;
        auto target = xlnt::uri(parser.attribute("Target"));

        if (target.path().is_absolute() && target_mode == xlnt::target_mode::internal)
        {
            target = uri(target.path().relative_to(path(part.string()).resolve(path("/"))).string());
        }

        relationships.emplace_back(parser.attribute("Id"),
            parser.attribute<xlnt::relationship_type>("Type"),
            xlnt::uri(part.string()), target, target_mode);

        expect_end_element(qn("relationships", "Relationship"));
    }

    expect_end_element(qn("relationships", "Relationships"));
    parser_ = nullptr;

    return relationships;
}

void xlsx_consumer::read_part(const std::vector<relationship> &rel_chain)
{
    const auto &manifest = target_.manifest();
    const auto part_path = manifest.canonicalize(rel_chain);
    auto part_streambuf = archive_->open(part_path);
    std::istream part_stream(part_streambuf.get());
    xml::parser parser(part_stream, part_path.string());
    parser_ = &parser;

    switch (rel_chain.back().type())
    {
    case relationship_type::core_properties:
        read_core_properties();
        break;

    case relationship_type::extended_properties:
        read_extended_properties();
        break;

    case relationship_type::custom_properties:
        read_custom_properties();
        break;

    case relationship_type::office_document:
        read_office_document(manifest.content_type(part_path));
        break;

    case relationship_type::connections:
        read_connections();
        break;

    case relationship_type::custom_xml_mappings:
        read_custom_xml_mappings();
        break;

    case relationship_type::external_workbook_references:
        read_external_workbook_references();
        break;

    case relationship_type::pivot_table:
        read_pivot_table();
        break;

    case relationship_type::shared_workbook_revision_headers:
        read_shared_workbook_revision_headers();
        break;

    case relationship_type::volatile_dependencies:
        read_volatile_dependencies();
        break;

    case relationship_type::shared_string_table:
        read_shared_string_table();
        break;

    case relationship_type::stylesheet:
        read_stylesheet();
        break;

    case relationship_type::theme:
        read_theme();
        break;

    case relationship_type::chartsheet:
        read_chartsheet(rel_chain.back().id());
        break;

    case relationship_type::dialogsheet:
        read_dialogsheet(rel_chain.back().id());
        break;

    case relationship_type::worksheet:
        read_worksheet(rel_chain.back().id());
        break;

    case relationship_type::thumbnail:
        read_image(part_path);
        break;

    case relationship_type::calculation_chain:
        read_calculation_chain();
        break;

    case relationship_type::hyperlink:
        break;

    case relationship_type::comments:
        break;

    case relationship_type::vml_drawing:
        break;

    case relationship_type::unknown:
        break;

    case relationship_type::printer_settings:
        read_binary(part_path);
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

    case relationship_type::vbaproject:
        read_binary(part_path);
        break;

    case relationship_type::image:
        read_image(part_path);
        break;
    }

    parser_ = nullptr;
}

void xlsx_consumer::populate_workbook(bool streaming)
{
    streaming_ = streaming;

    target_.clear();

    read_content_types();
    const auto root_path = path("/");

    for (const auto &package_rel : read_relationships(root_path))
    {
        manifest().register_relationship(package_rel);
    }

    for (auto package_rel : manifest().relationships(root_path))
    {
        if (package_rel.type() == relationship_type::office_document)
        {
            // Read the workbook after all the other package parts
            continue;
        }

        read_part({package_rel});
    }

    for (const auto &relationship_source_string : archive_->files())
    {
        for (const auto &part_rel : read_relationships(path(relationship_source_string)))
        {
            manifest().register_relationship(part_rel);
        }
    }

    read_part({manifest().relationship(root_path,
        relationship_type::office_document)});
}

// Package Parts

void xlsx_consumer::read_content_types()
{
    auto &manifest = target_.manifest();
    auto content_types_streambuf = archive_->open(path("[Content_Types].xml"));
    std::istream content_types_stream(content_types_streambuf.get());
    xml::parser parser(content_types_stream, "[Content_Types].xml");
    parser_ = &parser;

    expect_start_element(qn("content-types", "Types"), xml::content::complex);

    while (in_element(qn("content-types", "Types")))
    {
        auto current_element = expect_start_element(xml::content::complex);

        if (current_element == qn("content-types", "Default"))
        {
            auto extension = parser.attribute("Extension");
            auto content_type = parser.attribute("ContentType");
            manifest.register_default_type(extension, content_type);
        }
        else if (current_element == qn("content-types", "Override"))
        {
            auto part_name = parser.attribute("PartName");
            auto content_type = parser.attribute("ContentType");
            manifest.register_override_type(path(part_name), content_type);
        }
        else
        {
            unexpected_element(current_element);
        }

        expect_end_element(current_element);
    }

    expect_end_element(qn("content-types", "Types"));
}

void xlsx_consumer::read_core_properties()
{
    //qn("extended-properties", "Properties");
    //qn("custom-properties", "Properties");
    expect_start_element(qn("core-properties", "coreProperties"), xml::content::complex);

    while (in_element(qn("core-properties", "coreProperties")))
    {
        const auto property_element = expect_start_element(xml::content::simple);
        const auto prop = detail::from_string<core_property>(property_element.name());
        if (prop == core_property::created || prop == core_property::modified)
        {
            skip_attribute(qn("xsi", "type"));
        }
        target_.core_property(prop, read_text());
        expect_end_element(property_element);
    }

    expect_end_element(qn("core-properties", "coreProperties"));
}

void xlsx_consumer::read_extended_properties()
{
    expect_start_element(qn("extended-properties", "Properties"), xml::content::complex);

    while (in_element(qn("extended-properties", "Properties")))
    {
        const auto property_element = expect_start_element(xml::content::mixed);
        const auto prop = detail::from_string<extended_property>(property_element.name());
        target_.extended_property(prop, read_variant());
        expect_end_element(property_element);
    }

    expect_end_element(qn("extended-properties", "Properties"));
}

void xlsx_consumer::read_custom_properties()
{
    expect_start_element(qn("custom-properties", "Properties"), xml::content::complex);

    while (in_element(qn("custom-properties", "Properties")))
    {
        const auto property_element = expect_start_element(xml::content::complex);
        const auto prop = parser().attribute("name");
        const auto format_id = parser().attribute("fmtid");
        const auto property_id = parser().attribute("pid");
        target_.custom_property(prop, read_variant());
        expect_end_element(property_element);
    }

    expect_end_element(qn("custom-properties", "Properties"));
}

void xlsx_consumer::read_office_document(const std::string &content_type) // CT_Workbook
{
    if (content_type !=
            "application/vnd."
            "openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
        && content_type !=
            "application/vnd."
            "openxmlformats-officedocument.spreadsheetml.template.main+xml"
        && content_type !=
            "application/vnd."
            "ms-excel.sheet.macroEnabled.main+xml")
    {
        throw xlnt::invalid_file(content_type);
    }

    target_.d_->calculation_properties_.clear();

    expect_start_element(qn("workbook", "workbook"), xml::content::complex);
    skip_attribute(qn("mc", "Ignorable"));

    while (in_element(qn("workbook", "workbook")))
    {
        auto current_workbook_element = expect_start_element(xml::content::complex);

        if (current_workbook_element == qn("workbook", "fileVersion")) // CT_FileVersion 0-1
        {
            detail::workbook_impl::file_version_t file_version;

            if (parser().attribute_present("appName"))
            {
                file_version.app_name = parser().attribute("appName");
            }

            if (parser().attribute_present("lastEdited"))
            {
                file_version.last_edited = parser().attribute<std::size_t>("lastEdited");
            }

            if (parser().attribute_present("lowestEdited"))
            {
                file_version.lowest_edited = parser().attribute<std::size_t>("lowestEdited");
            }

            if (parser().attribute_present("lowestEdited"))
            {
                file_version.rup_build = parser().attribute<std::size_t>("rupBuild");
            }

            skip_attribute("codeName");

            target_.d_->file_version_ = file_version;
        }
        else if (current_workbook_element == qn("workbook", "fileSharing")) // CT_FileSharing 0-1
        {
            skip_remaining_content(current_workbook_element);
        }
        else if (current_workbook_element == qn("mc", "AlternateContent"))
        {
            while (in_element(qn("mc", "AlternateContent")))
            {
                auto alternate_content_element = expect_start_element(xml::content::complex);

                if (alternate_content_element == qn("mc", "Choice")
                    && parser().attribute_present("Requires")
                    && parser().attribute("Requires") == "x15")
                {
                    auto x15_element = expect_start_element(xml::content::simple);

                    if (x15_element == qn("x15ac", "absPath"))
                    {
                        target_.d_->abs_path_ = parser().attribute("url");
                    }

                    skip_remaining_content(x15_element);
                    expect_end_element(x15_element);
                }

                skip_remaining_content(alternate_content_element);
                expect_end_element(alternate_content_element);
            }
        }
        else if (current_workbook_element == qn("workbook", "workbookPr")) // CT_WorkbookPr 0-1
        {
            target_.base_date(parser().attribute_present("date1904") // optional, bool=false
                        && is_true(parser().attribute("date1904"))
                    ? calendar::mac_1904
                    : calendar::windows_1900);
            skip_attribute("showObjects"); // optional, ST_Objects="all"
            skip_attribute("showBorderUnselectedTables"); // optional, bool=true
            skip_attribute("filterPrivacy"); // optional, bool=false
            skip_attribute("promptedSolutions"); // optional, bool=false
            skip_attribute("showInkAnnotation"); // optional, bool=true
            skip_attribute("backupFile"); // optional, bool=false
            skip_attribute("saveExternalLinkValues"); // optional, bool=true
            skip_attribute("updateLinks"); // optional, ST_UpdateLinks="userSet"
            skip_attribute("codeName"); // optional, string
            skip_attribute("hidePivotFieldList"); // optional, bool=false
            skip_attribute("showPivotChartFilter"); // optional, bool=false
            skip_attribute("allowRefreshQuery"); // optional, bool=false
            skip_attribute("publishItems"); // optional, bool=false
            skip_attribute("checkCompatibility"); // optional, bool=false
            skip_attribute("autoCompressPictures"); // optional, bool=true
            skip_attribute("refreshAllConnections"); // optional, bool=false
            skip_attribute("defaultThemeVersion"); // optional, uint
            skip_attribute("dateCompatibility"); // optional, bool (undocumented)
        }
        else if (current_workbook_element == qn("workbook", "workbookProtection")) // CT_WorkbookProtection 0-1
        {
            skip_remaining_content(current_workbook_element);
        }
        else if (current_workbook_element == qn("workbook", "bookViews")) // CT_BookViews 0-1
        {
            while (in_element(qn("workbook", "bookViews")))
            {
                expect_start_element(qn("workbook", "workbookView"), xml::content::simple);
                skip_attributes({"firstSheet", "showHorizontalScroll",
                    "showSheetTabs", "showVerticalScroll"});

                workbook_view view;

                if (parser().attribute_present("xWindow"))
                {
                    view.x_window = parser().attribute<int>("xWindow");
                }

                if (parser().attribute_present("yWindow"))
                {
                    view.y_window = parser().attribute<int>("yWindow");
                }

                if (parser().attribute_present("windowWidth"))
                {
                    view.window_width = parser().attribute<std::size_t>("windowWidth");
                }

                if (parser().attribute_present("windowHeight"))
                {
                    view.window_height = parser().attribute<std::size_t>("windowHeight");
                }

                if (parser().attribute_present("tabRatio"))
                {
                    view.tab_ratio = parser().attribute<std::size_t>("tabRatio");
                }

                if (parser().attribute_present("activeTab"))
                {
                    view.active_tab = parser().attribute<std::size_t>("activeTab");
                    target_.d_->active_sheet_index_.set(view.active_tab.get());
                }

                target_.view(view);

                skip_attributes();
                expect_end_element(qn("workbook", "workbookView"));
            }
        }
        else if (current_workbook_element == qn("workbook", "sheets")) // CT_Sheets 1
        {
            std::size_t index = 0;

            while (in_element(qn("workbook", "sheets")))
            {
                expect_start_element(qn("spreadsheetml", "sheet"), xml::content::simple);

                auto title = parser().attribute("name");

                sheet_title_index_map_[title] = index++;
                sheet_title_id_map_[title] = parser().attribute<std::size_t>("sheetId");
                target_.d_->sheet_title_rel_id_map_[title] = parser().attribute(qn("r", "id"));

                bool hidden = parser().attribute<std::string>("state", "") == "hidden";
                target_.d_->sheet_hidden_.push_back(hidden);

                expect_end_element(qn("spreadsheetml", "sheet"));
            }
        }
        else if (current_workbook_element == qn("workbook", "functionGroups")) // CT_FunctionGroups 0-1
        {
            skip_remaining_content(current_workbook_element);
        }
        else if (current_workbook_element == qn("workbook", "externalReferences")) // CT_ExternalReferences 0-1
        {
            skip_remaining_content(current_workbook_element);
        }
        else if (current_workbook_element == qn("workbook", "definedNames")) // CT_DefinedNames 0-1
        {
            while (in_element(qn("workbook", "definedNames")))
            {
                expect_start_element(qn("spreadsheetml", "definedName"), xml::content::mixed);

                defined_name name;
                name.name = parser().attribute("name");
                name.sheet_id = parser().attribute<std::size_t>("localSheetId");
                name.hidden = false;
                if (parser().attribute_present("hidden"))
                {
                    name.hidden = is_true(parser().attribute("hidden"));
                }
                parser().attribute_map(); // skip remaining attributes
                name.value = read_text();
                defined_names_.push_back(name);
                
                expect_end_element(qn("spreadsheetml", "definedName"));
            }
        }
        else if (current_workbook_element == qn("workbook", "calcPr")) // CT_CalcPr 0-1
        {
            xlnt::calculation_properties calc_props;
            if (parser().attribute_present("calcId"))
            {
                calc_props.calc_id = parser().attribute<std::size_t>("calcId");
            }
            if (parser().attribute_present("concurrentCalc"))
            {
                calc_props.concurrent_calc = is_true(parser().attribute("concurrentCalc"));
            }
            target_.calculation_properties(calc_props);
            parser().attribute_map(); // skip remaining
        }
        else if (current_workbook_element == qn("workbook", "oleSize")) // CT_OleSize 0-1
        {
            skip_remaining_content(current_workbook_element);
        }
        else if (current_workbook_element == qn("workbook", "customWorkbookViews")) // CT_CustomWorkbookViews 0-1
        {
            skip_remaining_content(current_workbook_element);
        }
        else if (current_workbook_element == qn("workbook", "pivotCaches")) // CT_PivotCaches 0-1
        {
            skip_remaining_content(current_workbook_element);
        }
        else if (current_workbook_element == qn("workbook", "smartTagPr")) // CT_SmartTagPr 0-1
        {
            skip_remaining_content(current_workbook_element);
        }
        else if (current_workbook_element == qn("workbook", "smartTagTypes")) // CT_SmartTagTypes 0-1
        {
            skip_remaining_content(current_workbook_element);
        }
        else if (current_workbook_element == qn("workbook", "webPublishing")) // CT_WebPublishing 0-1
        {
            skip_remaining_content(current_workbook_element);
        }
        else if (current_workbook_element == qn("workbook", "fileRecoveryPr")) // CT_FileRecoveryPr 0+
        {
            skip_remaining_content(current_workbook_element);
        }
        else if (current_workbook_element == qn("workbook", "webPublishObjects")) // CT_WebPublishObjects 0-1
        {
            skip_remaining_content(current_workbook_element);
        }
        else if (current_workbook_element == qn("workbook", "extLst")) // CT_ExtensionList 0-1
        {
            while (in_element(qn("workbook", "extLst")))
            {
                auto extension_element = expect_start_element(xml::content::complex);

                if (extension_element == qn("workbook", "ext")
                    && parser().attribute_present("uri")
                    && parser().attribute("uri") == "{7523E5D3-25F3-A5E0-1632-64F254C22452}")
                {
                    auto arch_id_extension_element = expect_start_element(xml::content::simple);

                    if (arch_id_extension_element == qn("mx", "ArchID"))
                    {
                        target_.d_->arch_id_flags_ = parser().attribute<std::size_t>("Flags");
                    }

                    skip_remaining_content(arch_id_extension_element);
                    expect_end_element(arch_id_extension_element);
                }

                skip_remaining_content(extension_element);
                expect_end_element(extension_element);
            }
        }
        else
        {
            unexpected_element(current_workbook_element);
        }

        expect_end_element(current_workbook_element);
    }

    expect_end_element(qn("workbook", "workbook"));

    auto workbook_rel = manifest().relationship(path("/"), relationship_type::office_document);
    auto workbook_path = workbook_rel.target().path();

    const auto rel_types = {
        relationship_type::shared_string_table,
        relationship_type::stylesheet,
        relationship_type::theme,
        relationship_type::vbaproject,
    };

    for (auto rel_type : rel_types)
    {
        if (manifest().has_relationship(workbook_path, rel_type))
        {
            read_part({workbook_rel,
                manifest().relationship(workbook_path, rel_type)});
        }
    }

    for (auto worksheet_rel : manifest().relationships(workbook_path, relationship_type::worksheet))
    {
        auto title = std::find_if(target_.d_->sheet_title_rel_id_map_.begin(),
            target_.d_->sheet_title_rel_id_map_.end(),
            [&](const std::pair<std::string, std::string> &p) {
                return p.second == worksheet_rel.id();
            })->first;

        auto id = sheet_title_id_map_[title];
        auto index = sheet_title_index_map_[title];

        auto insertion_iter = target_.d_->worksheets_.begin();
        while (insertion_iter != target_.d_->worksheets_.end()
            && sheet_title_index_map_[insertion_iter->title_] < index)
        {
            ++insertion_iter;
        }

        current_worksheet_ = &*target_.d_->worksheets_.emplace(insertion_iter, &target_, id, title);

        if (!streaming_)
        {
            read_part({workbook_rel, worksheet_rel});
        }
    }
}

// Write Workbook Relationship Target Parts

void xlsx_consumer::read_calculation_chain()
{
}

void xlsx_consumer::read_chartsheet(const std::string & /*title*/)
{
}

void xlsx_consumer::read_connections()
{
}

void xlsx_consumer::read_custom_property()
{
}

void xlsx_consumer::read_custom_xml_mappings()
{
}

void xlsx_consumer::read_dialogsheet(const std::string & /*title*/)
{
}

void xlsx_consumer::read_external_workbook_references()
{
}

void xlsx_consumer::read_pivot_table()
{
}

void xlsx_consumer::read_shared_string_table()
{
    expect_start_element(qn("spreadsheetml", "sst"), xml::content::complex);
    skip_attributes({"count"});

    bool has_unique_count = false;
    std::size_t unique_count = 0;

    if (parser().attribute_present("uniqueCount"))
    {
        has_unique_count = true;
        unique_count = parser().attribute<std::size_t>("uniqueCount");
    }

    while (in_element(qn("spreadsheetml", "sst")))
    {
        expect_start_element(qn("spreadsheetml", "si"), xml::content::complex);
        auto rt = read_rich_text(qn("spreadsheetml", "si"));
        target_.add_shared_string(rt, true);
        expect_end_element(qn("spreadsheetml", "si"));
    }

    expect_end_element(qn("spreadsheetml", "sst"));

    if (has_unique_count && unique_count != target_.shared_strings().size())
    {
        throw invalid_file("sizes don't match");
    }
}

void xlsx_consumer::read_shared_workbook_revision_headers()
{
}

void xlsx_consumer::read_shared_workbook()
{
}

void xlsx_consumer::read_shared_workbook_user_data()
{
}

void xlsx_consumer::read_stylesheet()
{
    target_.impl().stylesheet_ = detail::stylesheet();
    auto &stylesheet = target_.impl().stylesheet_.get();

    expect_start_element(qn("spreadsheetml", "styleSheet"), xml::content::complex);
    skip_attributes({qn("mc", "Ignorable")});

    std::vector<std::pair<style_impl, std::size_t>> styles;
    std::vector<std::pair<format_impl, std::size_t>> format_records;
    std::vector<std::pair<format_impl, std::size_t>> style_records;

    while (in_element(qn("spreadsheetml", "styleSheet")))
    {
        auto current_style_element = expect_start_element(xml::content::complex);

        if (current_style_element == qn("spreadsheetml", "borders"))
        {
            auto &borders = stylesheet.borders;
            auto count = parser().attribute<std::size_t>("count");

            while (in_element(qn("spreadsheetml", "borders")))
            {
                borders.push_back(xlnt::border());
                auto &border = borders.back();

                expect_start_element(qn("spreadsheetml", "border"), xml::content::complex);

                auto diagonal = diagonal_direction::neither;

                if (parser().attribute_present("diagonalDown") && parser().attribute("diagonalDown") == "1")
                {
                    diagonal = diagonal_direction::down;
                }

                if (parser().attribute_present("diagonalUp") && parser().attribute("diagonalUp") == "1")
                {
                    diagonal = diagonal == diagonal_direction::down ? diagonal_direction::both : diagonal_direction::up;
                }

                if (diagonal != diagonal_direction::neither)
                {
                    border.diagonal(diagonal);
                }

                while (in_element(qn("spreadsheetml", "border")))
                {
                    auto current_side_element = expect_start_element(xml::content::complex);

                    xlnt::border::border_property side;

                    if (parser().attribute_present("style"))
                    {
                        side.style(parser().attribute<xlnt::border_style>("style"));
                    }

                    if (in_element(current_side_element))
                    {
                        expect_start_element(qn("spreadsheetml", "color"), xml::content::complex);
                        side.color(read_color());
                        expect_end_element(qn("spreadsheetml", "color"));
                    }

                    expect_end_element(current_side_element);

                    auto side_type = xml::value_traits<xlnt::border_side>::parse(current_side_element.name(), parser());
                    border.side(side_type, side);
                }

                expect_end_element(qn("spreadsheetml", "border"));
            }

            if (count != borders.size())
            {
                throw xlnt::exception("border counts don't match");
            }
        }
        else if (current_style_element == qn("spreadsheetml", "fills"))
        {
            auto &fills = stylesheet.fills;
            auto count = parser().attribute<std::size_t>("count");

            while (in_element(qn("spreadsheetml", "fills")))
            {
                fills.push_back(xlnt::fill());
                auto &new_fill = fills.back();

                expect_start_element(qn("spreadsheetml", "fill"), xml::content::complex);
                auto fill_element = expect_start_element(xml::content::complex);

                if (fill_element == qn("spreadsheetml", "patternFill"))
                {
                    xlnt::pattern_fill pattern;

                    if (parser().attribute_present("patternType"))
                    {
                        pattern.type(parser().attribute<xlnt::pattern_fill_type>("patternType"));

                        while (in_element(qn("spreadsheetml", "patternFill")))
                        {
                            auto pattern_type_element = expect_start_element(xml::content::complex);

                            if (pattern_type_element == qn("spreadsheetml", "fgColor"))
                            {
                                pattern.foreground(read_color());
                            }
                            else if (pattern_type_element == qn("spreadsheetml", "bgColor"))
                            {
                                pattern.background(read_color());
                            }
                            else
                            {
                                unexpected_element(pattern_type_element);
                            }

                            expect_end_element(pattern_type_element);
                        }
                    }

                    new_fill = pattern;
                }
                else if (fill_element == qn("spreadsheetml", "gradientFill"))
                {
                    xlnt::gradient_fill gradient;

                    if (parser().attribute_present("type"))
                    {
                        gradient.type(parser().attribute<xlnt::gradient_fill_type>("type"));
                    }
                    else
                    {
                        gradient.type(xlnt::gradient_fill_type::linear);
                    }

                    while (in_element(qn("spreadsheetml", "gradientFill")))
                    {
                        expect_start_element(qn("spreadsheetml", "stop"), xml::content::complex);
                        auto position = converter_.deserialise(parser().attribute("position"));
                        expect_start_element(qn("spreadsheetml", "color"), xml::content::complex);
                        auto color = read_color();
                        expect_end_element(qn("spreadsheetml", "color"));
                        expect_end_element(qn("spreadsheetml", "stop"));

                        gradient.add_stop(position, color);
                    }

                    new_fill = gradient;
                }
                else
                {
                    unexpected_element(fill_element);
                }

                expect_end_element(fill_element);
                expect_end_element(qn("spreadsheetml", "fill"));
            }

            if (count != fills.size())
            {
                throw xlnt::exception("counts don't match");
            }
        }
        else if (current_style_element == qn("spreadsheetml", "fonts"))
        {
            auto &fonts = stylesheet.fonts;
            auto count = parser().attribute<std::size_t>("count", 0);

            if (parser().attribute_present(qn("x14ac", "knownFonts")))
            {
                target_.enable_known_fonts();
            }

            while (in_element(qn("spreadsheetml", "fonts")))
            {
                fonts.push_back(xlnt::font());
                auto &new_font = stylesheet.fonts.back();

                expect_start_element(qn("spreadsheetml", "font"), xml::content::complex);

                while (in_element(qn("spreadsheetml", "font")))
                {
                    auto font_property_element = expect_start_element(xml::content::simple);

                    if (font_property_element == qn("spreadsheetml", "sz"))
                    {
                        new_font.size(converter_.deserialise(parser().attribute("val")));
                    }
                    else if (font_property_element == qn("spreadsheetml", "name"))
                    {
                        new_font.name(parser().attribute("val"));
                    }
                    else if (font_property_element == qn("spreadsheetml", "color"))
                    {
                        new_font.color(read_color());
                    }
                    else if (font_property_element == qn("spreadsheetml", "family"))
                    {
                        new_font.family(parser().attribute<std::size_t>("val"));
                    }
                    else if (font_property_element == qn("spreadsheetml", "scheme"))
                    {
                        new_font.scheme(parser().attribute("val"));
                    }
                    else if (font_property_element == qn("spreadsheetml", "b"))
                    {
                        if (parser().attribute_present("val"))
                        {
                            new_font.bold(is_true(parser().attribute("val")));
                        }
                        else
                        {
                            new_font.bold(true);
                        }
                    }
                    else if (font_property_element == qn("spreadsheetml", "vertAlign"))
                    {
                        auto vert_align = parser().attribute("val");

                        if (vert_align == "superscript")
                        {
                            new_font.superscript(true);
                        }
                        else if (vert_align == "subscript")
                        {
                            new_font.subscript(true);
                        }
                    }
                    else if (font_property_element == qn("spreadsheetml", "strike"))
                    {
                        if (parser().attribute_present("val"))
                        {
                            new_font.strikethrough(is_true(parser().attribute("val")));
                        }
                        else
                        {
                            new_font.strikethrough(true);
                        }
                    }
                    else if (font_property_element == qn("spreadsheetml", "outline"))
                    {
                        if (parser().attribute_present("val"))
                        {
                            new_font.outline(is_true(parser().attribute("val")));
                        }
                        else
                        {
                            new_font.outline(true);
                        }
                    }
                    else if (font_property_element == qn("spreadsheetml", "shadow"))
                    {
                        if (parser().attribute_present("val"))
                        {
                            new_font.shadow(is_true(parser().attribute("val")));
                        }
                        else
                        {
                            new_font.shadow(true);
                        }
                    }
                    else if (font_property_element == qn("spreadsheetml", "i"))
                    {
                        if (parser().attribute_present("val"))
                        {
                            new_font.italic(is_true(parser().attribute("val")));
                        }
                        else
                        {
                            new_font.italic(true);
                        }
                    }
                    else if (font_property_element == qn("spreadsheetml", "u"))
                    {
                        if (parser().attribute_present("val"))
                        {
                            new_font.underline(parser().attribute<xlnt::font::underline_style>("val"));
                        }
                        else
                        {
                            new_font.underline(xlnt::font::underline_style::single);
                        }
                    }
                    else if (font_property_element == qn("spreadsheetml", "charset"))
                    {
                        if (parser().attribute_present("val"))
                        {
                            parser().attribute("val");
                        }
                    }
                    else
                    {
                        unexpected_element(font_property_element);
                    }

                    expect_end_element(font_property_element);
                }

                expect_end_element(qn("spreadsheetml", "font"));
            }

            if (count != stylesheet.fonts.size())
            {
                // throw xlnt::exception("counts don't match");
            }
        }
        else if (current_style_element == qn("spreadsheetml", "numFmts"))
        {
            auto &number_formats = stylesheet.number_formats;
            auto count = parser().attribute<std::size_t>("count");

            while (in_element(qn("spreadsheetml", "numFmts")))
            {
                expect_start_element(qn("spreadsheetml", "numFmt"), xml::content::simple);

                auto format_string = parser().attribute("formatCode");

                if (format_string == "GENERAL")
                {
                    format_string = "General";
                }

                xlnt::number_format nf;

                nf.format_string(format_string);
                nf.id(parser().attribute<std::size_t>("numFmtId"));

                expect_end_element(qn("spreadsheetml", "numFmt"));

                number_formats.push_back(nf);
            }

            if (count != number_formats.size())
            {
                throw xlnt::exception("counts don't match");
            }
        }
        else if (current_style_element == qn("spreadsheetml", "cellStyles"))
        {
            auto count = parser().attribute<std::size_t>("count");

            while (in_element(qn("spreadsheetml", "cellStyles")))
            {
                auto &data = *styles.emplace(styles.end());

                expect_start_element(qn("spreadsheetml", "cellStyle"), xml::content::simple);

                data.first.name = parser().attribute("name");
                data.second = parser().attribute<std::size_t>("xfId");

                if (parser().attribute_present("builtinId"))
                {
                    data.first.builtin_id = parser().attribute<std::size_t>("builtinId");
                }

                if (parser().attribute_present("hidden"))
                {
                    data.first.hidden_style = is_true(parser().attribute("hidden"));
                }

                if (parser().attribute_present("customBuiltin"))
                {
                    data.first.custom_builtin = is_true(parser().attribute("customBuiltin"));
                }

                expect_end_element(qn("spreadsheetml", "cellStyle"));
            }

            if (count != styles.size())
            {
                throw xlnt::exception("counts don't match");
            }
        }
        else if (current_style_element == qn("spreadsheetml", "cellStyleXfs")
            || current_style_element == qn("spreadsheetml", "cellXfs"))
        {
            auto in_style_records = current_style_element.name() == "cellStyleXfs";
            auto count = parser().attribute<std::size_t>("count");

            while (in_element(current_style_element))
            {
                expect_start_element(qn("spreadsheetml", "xf"), xml::content::complex);

                auto &record = *(!in_style_records
                        ? format_records.emplace(format_records.end())
                        : style_records.emplace(style_records.end()));

                if (parser().attribute_present("applyBorder"))
                {
                    record.first.border_applied = is_true(parser().attribute("applyBorder"));
                }
                record.first.border_id = parser().attribute_present("borderId")
                    ? parser().attribute<std::size_t>("borderId")
                    : optional<std::size_t>();

                if (parser().attribute_present("applyFill"))
                {
                    record.first.fill_applied = is_true(parser().attribute("applyFill"));
                }
                record.first.fill_id = parser().attribute_present("fillId")
                    ? parser().attribute<std::size_t>("fillId")
                    : optional<std::size_t>();

                if (parser().attribute_present("applyFont"))
                {
                    record.first.font_applied = is_true(parser().attribute("applyFont"));
                }
                record.first.font_id = parser().attribute_present("fontId")
                    ? parser().attribute<std::size_t>("fontId")
                    : optional<std::size_t>();

                if (parser().attribute_present("applyNumberFormat"))
                {
                    record.first.number_format_applied = is_true(parser().attribute("applyNumberFormat"));
                }
                record.first.number_format_id = parser().attribute_present("numFmtId")
                    ? parser().attribute<std::size_t>("numFmtId")
                    : optional<std::size_t>();

                auto apply_alignment_present = parser().attribute_present("applyAlignment");
                if (apply_alignment_present)
                {
                    record.first.alignment_applied = is_true(parser().attribute("applyAlignment"));
                }

                auto apply_protection_present = parser().attribute_present("applyProtection");
                if (apply_protection_present)
                {
                    record.first.protection_applied = is_true(parser().attribute("applyProtection"));
                }

                record.first.pivot_button_ = parser().attribute_present("pivotButton")
                    && is_true(parser().attribute("pivotButton"));
                record.first.quote_prefix_ = parser().attribute_present("quotePrefix")
                    && is_true(parser().attribute("quotePrefix"));

                if (parser().attribute_present("xfId"))
                {
                    record.second = parser().attribute<std::size_t>("xfId");
                }

                while (in_element(qn("spreadsheetml", "xf")))
                {
                    auto xf_child_element = expect_start_element(xml::content::simple);

                    if (xf_child_element == qn("spreadsheetml", "alignment"))
                    {
                        record.first.alignment_id = stylesheet.alignments.size();
                        auto &alignment = *stylesheet.alignments.emplace(stylesheet.alignments.end());

                        if (parser().attribute_present("wrapText"))
                        {
                            alignment.wrap(is_true(parser().attribute("wrapText")));
                        }

                        if (parser().attribute_present("shrinkToFit"))
                        {
                            alignment.shrink(is_true(parser().attribute("shrinkToFit")));
                        }

                        if (parser().attribute_present("indent"))
                        {
                            alignment.indent(parser().attribute<int>("indent"));
                        }

                        if (parser().attribute_present("textRotation"))
                        {
                            alignment.rotation(parser().attribute<int>("textRotation"));
                        }

                        if (parser().attribute_present("vertical"))
                        {
                            alignment.vertical(parser().attribute<xlnt::vertical_alignment>("vertical"));
                        }

                        if (parser().attribute_present("horizontal"))
                        {
                            alignment.horizontal(parser().attribute<xlnt::horizontal_alignment>("horizontal"));
                        }

                        if (parser().attribute_present("readingOrder"))
                        {
                            parser().attribute<int>("readingOrder");
                        }
                    }
                    else if (xf_child_element == qn("spreadsheetml", "protection"))
                    {
                        record.first.protection_id = stylesheet.protections.size();
                        auto &protection = *stylesheet.protections.emplace(stylesheet.protections.end());

                        protection.locked(parser().attribute_present("locked")
                            && is_true(parser().attribute("locked")));
                        protection.hidden(parser().attribute_present("hidden")
                            && is_true(parser().attribute("hidden")));
                    }
                    else
                    {
                        unexpected_element(xf_child_element);
                    }

                    expect_end_element(xf_child_element);
                }

                expect_end_element(qn("spreadsheetml", "xf"));
            }

            if ((in_style_records && count != style_records.size())
                || (!in_style_records && count != format_records.size()))
            {
                throw xlnt::exception("counts don't match");
            }
        }
        else if (current_style_element == qn("spreadsheetml", "dxfs"))
        {
            auto count = parser().attribute<std::size_t>("count");
            std::size_t processed = 0;

            while (in_element(current_style_element))
            {
                auto current_element = expect_start_element(xml::content::mixed);
                skip_remaining_content(current_element);
                expect_end_element(current_element);
                ++processed;
            }

            if (count != processed)
            {
                throw xlnt::exception("counts don't match");
            }
        }
        else if (current_style_element == qn("spreadsheetml", "tableStyles"))
        {
            skip_attribute("defaultTableStyle");
            skip_attribute("defaultPivotStyle");

            auto count = parser().attribute<std::size_t>("count");
            std::size_t processed = 0;

            while (in_element(qn("spreadsheetml", "tableStyles")))
            {
                auto current_element = expect_start_element(xml::content::complex);
                skip_remaining_content(current_element);
                expect_end_element(current_element);
                ++processed;
            }

            if (count != processed)
            {
                throw xlnt::exception("counts don't match");
            }
        }
        else if (current_style_element == qn("spreadsheetml", "extLst"))
        {
            while (in_element(qn("spreadsheetml", "extLst")))
            {
                expect_start_element(qn("spreadsheetml", "ext"), xml::content::complex);

                const auto uri = parser().attribute("uri");

                if (uri == "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}") // slicerStyles
                {
                    expect_start_element(qn("x14", "slicerStyles"), xml::content::simple);
                    stylesheet.default_slicer_style = parser().attribute("defaultSlicerStyle");
                    expect_end_element(qn("x14", "slicerStyles"));
                }
                else
                {
                    skip_remaining_content(qn("spreadsheetml", "ext"));
                }

                expect_end_element(qn("spreadsheetml", "ext"));
            }
        }
        else if (current_style_element == qn("spreadsheetml", "colors")) // CT_Colors 0-1
        {
            while (in_element(qn("spreadsheetml", "colors")))
            {
                auto colors_child_element = expect_start_element(xml::content::complex);

                if (colors_child_element == qn("spreadsheetml", "indexedColors")) // CT_IndexedColors 0-1
                {
                    while (in_element(colors_child_element))
                    {
                        expect_start_element(qn("spreadsheetml", "rgbColor"), xml::content::simple);
                        stylesheet.colors.push_back(read_color());
                        expect_end_element(qn("spreadsheetml", "rgbColor"));
                    }
                }
                else if (colors_child_element == qn("spreadsheetml", "mruColors")) // CT_MRUColors
                {
                    skip_remaining_content(colors_child_element);
                }
                else
                {
                    unexpected_element(colors_child_element);
                }

                expect_end_element(colors_child_element);
            }
        }
        else
        {
            unexpected_element(current_style_element);
        }

        expect_end_element(current_style_element);
    }

    expect_end_element(qn("spreadsheetml", "styleSheet"));

    std::size_t xf_id = 0;

    for (const auto &record : style_records)
    {
        auto style_iter = std::find_if(styles.begin(), styles.end(),
            [&xf_id](const std::pair<style_impl, std::size_t> &s) { return s.second == xf_id; });
        ++xf_id;

        if (style_iter == styles.end()) continue;

        auto new_style = stylesheet.create_style(style_iter->first.name);

        new_style.d_->pivot_button_ = style_iter->first.pivot_button_;
        new_style.d_->quote_prefix_ = style_iter->first.quote_prefix_;
        new_style.d_->formatting_record_id = style_iter->first.formatting_record_id;
        new_style.d_->hidden_style = style_iter->first.hidden_style;
        new_style.d_->custom_builtin = style_iter->first.custom_builtin;
        new_style.d_->hidden_style = style_iter->first.hidden_style;
        new_style.d_->builtin_id = style_iter->first.builtin_id;
        new_style.d_->outline_style = style_iter->first.outline_style;

        new_style.d_->alignment_applied = record.first.alignment_applied;
        new_style.d_->alignment_id = record.first.alignment_id;
        new_style.d_->border_applied = record.first.border_applied;
        new_style.d_->border_id = record.first.border_id;
        new_style.d_->fill_applied = record.first.fill_applied;
        new_style.d_->fill_id = record.first.fill_id;
        new_style.d_->font_applied = record.first.font_applied;
        new_style.d_->font_id = record.first.font_id;
        new_style.d_->number_format_applied = record.first.number_format_applied;
        new_style.d_->number_format_id = record.first.number_format_id;
    }

    std::size_t record_index = 0;

    for (const auto &record : format_records)
    {
        stylesheet.format_impls.push_back(format_impl());
        auto &new_format = stylesheet.format_impls.back();

        new_format.id = record_index++;
        new_format.parent = &stylesheet;

        ++new_format.references;

        new_format.alignment_id = record.first.alignment_id;
        new_format.alignment_applied = record.first.alignment_applied;
        new_format.border_id = record.first.border_id;
        new_format.border_applied = record.first.border_applied;
        new_format.fill_id = record.first.fill_id;
        new_format.fill_applied = record.first.fill_applied;
        new_format.font_id = record.first.font_id;
        new_format.font_applied = record.first.font_applied;
        new_format.number_format_id = record.first.number_format_id;
        new_format.number_format_applied = record.first.number_format_applied;
        new_format.protection_id = record.first.protection_id;
        new_format.protection_applied = record.first.protection_applied;
        new_format.pivot_button_ = record.first.pivot_button_;
        new_format.quote_prefix_ = record.first.quote_prefix_;

        set_style_by_xfid(styles, record.second, new_format.style);
    }
}

void xlsx_consumer::read_theme()
{
    auto workbook_rel = manifest().relationship(path("/"),
        relationship_type::office_document);
    auto theme_rel = manifest().relationship(workbook_rel.target().path(),
        relationship_type::theme);
    auto theme_path = manifest().canonicalize({workbook_rel, theme_rel});

    target_.theme(theme());

    if (manifest().has_relationship(theme_path, relationship_type::image))
    {
        read_part({workbook_rel, theme_rel,
            manifest().relationship(theme_path,
                relationship_type::image)});
    }
}

void xlsx_consumer::read_volatile_dependencies()
{
}

// Sheet Relationship Target Parts

void xlsx_consumer::read_vml_drawings(worksheet /*ws*/)
{
}

void xlsx_consumer::read_comments(worksheet ws)
{
    std::vector<std::string> authors;

    expect_start_element(qn("spreadsheetml", "comments"), xml::content::complex);
    // name space can be ignored
    skip_attribute(qn("mc", "Ignorable"));
    expect_start_element(qn("spreadsheetml", "authors"), xml::content::complex);

    while (in_element(qn("spreadsheetml", "authors")))
    {
        expect_start_element(qn("spreadsheetml", "author"), xml::content::simple);
        authors.push_back(read_text());
        expect_end_element(qn("spreadsheetml", "author"));
    }

    expect_end_element(qn("spreadsheetml", "authors"));
    expect_start_element(qn("spreadsheetml", "commentList"), xml::content::complex);

    while (in_element(xml::qname(qn("spreadsheetml", "commentList"))))
    {
        expect_start_element(qn("spreadsheetml", "comment"), xml::content::complex);

        skip_attribute("shapeId");
        auto cell_ref = parser().attribute("ref");
        auto author_id = parser().attribute<std::size_t>("authorId");

        expect_start_element(qn("spreadsheetml", "text"), xml::content::complex);

        ws.cell(cell_ref).comment(comment(read_rich_text(qn("spreadsheetml", "text")), authors.at(author_id)));

        expect_end_element(qn("spreadsheetml", "text"));

        if (in_element(xml::qname(qn("spreadsheetml", "comment"))))
        {
            expect_start_element(qn("mc", "AlternateContent"), xml::content::complex);
            skip_remaining_content(qn("mc", "AlternateContent"));
            expect_end_element(qn("mc", "AlternateContent"));
        }

        expect_end_element(qn("spreadsheetml", "comment"));
    }

    expect_end_element(qn("spreadsheetml", "commentList"));
    expect_end_element(qn("spreadsheetml", "comments"));
}

void xlsx_consumer::read_drawings(worksheet ws, const path &part)
{
    auto images = manifest().relationships(part, relationship_type::image);

    auto sd = drawing::spreadsheet_drawing(parser());

    for (const auto &image_rel_id : sd.get_embed_ids())
    {
        auto image_rel = std::find_if(images.begin(), images.end(),
            [&](const relationship &r) { return r.id() == image_rel_id; });

        if (image_rel != images.end())
        {
            const auto url = image_rel->target().path().resolve(part.parent());

            read_image(url);
        }
    }

    ws.d_->drawing_ = sd;
}

// Unknown Parts

void xlsx_consumer::read_unknown_parts()
{
}

void xlsx_consumer::read_unknown_relationships()
{
}

void xlsx_consumer::read_image(const xlnt::path &image_path)
{
    auto image_streambuf = archive_->open(image_path);
    vector_ostreambuf buffer(target_.d_->images_[image_path.string()]);
    std::ostream out_stream(&buffer);
    out_stream << image_streambuf.get();
}

void xlsx_consumer::read_binary(const xlnt::path &binary_path)
{
    auto binary_streambuf = archive_->open(binary_path);
    vector_ostreambuf buffer(target_.d_->binaries_[binary_path.string()]);
    std::ostream out_stream(&buffer);
    out_stream << binary_streambuf.get();
}

std::string xlsx_consumer::read_text()
{
    auto text = std::string();

    while (parser().peek() == xml::parser::event_type::characters)
    {
        parser().next_expect(xml::parser::event_type::characters);
        text.append(parser().value());
    }

    return text;
}

variant xlsx_consumer::read_variant()
{
    auto value = variant(read_text());

    if (in_element(stack_.back()))
    {
        auto element = expect_start_element(xml::content::mixed);
        auto text = read_text();

        if (element == qn("vt", "lpwstr") || element == qn("vt", "lpstr"))
        {
            value = variant(text);
        }
        if (element == qn("vt", "i4"))
        {
            value = variant(std::stoi(text));
        }
        if (element == qn("vt", "bool"))
        {
            value = variant(is_true(text));
        }
        else if (element == qn("vt", "vector"))
        {
            auto size = parser().attribute<std::size_t>("size");
            auto base_type = parser().attribute("baseType");

            std::vector<variant> vector;

            for (auto i = std::size_t(0); i < size; ++i)
            {
                if (base_type == "variant")
                {
                    expect_start_element(qn("vt", "variant"), xml::content::complex);
                }

                vector.push_back(read_variant());

                if (base_type == "variant")
                {
                    expect_end_element(qn("vt", "variant"));
                    read_text();
                }
            }

            value = variant(vector);
        }

        expect_end_element(element);
        read_text();
    }

    return value;
}

void xlsx_consumer::skip_attributes(const std::vector<std::string> &names)
{
    for (const auto &name : names)
    {
        if (parser().attribute_present(name))
        {
            parser().attribute(name);
        }
    }
}

void xlsx_consumer::skip_attributes(const std::vector<xml::qname> &names)
{
    for (const auto &name : names)
    {
        if (parser().attribute_present(name))
        {
            parser().attribute(name);
        }
    }
}

void xlsx_consumer::skip_attributes()
{
    parser().attribute_map();
}

void xlsx_consumer::skip_attribute(const xml::qname &name)
{
    if (parser().attribute_present(name))
    {
        parser().attribute(name);
    }
}

void xlsx_consumer::skip_attribute(const std::string &name)
{
    if (parser().attribute_present(name))
    {
        parser().attribute(name);
    }
}

void xlsx_consumer::skip_remaining_content(const xml::qname &name)
{
    // start by assuming we've already parsed the opening tag

    skip_attributes();
    read_text();

    // continue until the closing tag is reached
    while (in_element(name))
    {
        auto child_element = expect_start_element(xml::content::mixed);
        skip_remaining_content(child_element);
        expect_end_element(child_element);
        read_text(); // trailing character content (usually whitespace)
    }
}

bool xlsx_consumer::in_element(const xml::qname &name)
{
    return parser().peek() != xml::parser::event_type::end_element
        && stack_.back() == name;
}

xml::qname xlsx_consumer::expect_start_element(xml::content content)
{
    parser().next_expect(xml::parser::event_type::start_element);
    parser().content(content);
    stack_.push_back(parser().qname());

    const auto xml_space = qn("xml", "space");
    preserve_space_ = parser().attribute_present(xml_space) ? parser().attribute(xml_space) == "preserve" : false;

    return stack_.back();
}

void xlsx_consumer::expect_start_element(const xml::qname &name, xml::content content)
{
    parser().next_expect(xml::parser::event_type::start_element, name);
    parser().content(content);
    stack_.push_back(name);

    const auto xml_space = qn("xml", "space");
    preserve_space_ = parser().attribute_present(xml_space) ? parser().attribute(xml_space) == "preserve" : false;
}

void xlsx_consumer::expect_end_element(const xml::qname &name)
{
    parser().attribute_map();
    parser().next_expect(xml::parser::event_type::end_element, name);
    stack_.pop_back();
}

void xlsx_consumer::unexpected_element(const xml::qname &name)
{
#ifdef THROW_ON_INVALID_XML
    throw xlnt::exception(name.string());
#else
    skip_remaining_content(name);
#endif
}

rich_text xlsx_consumer::read_rich_text(const xml::qname &parent)
{
    const auto &xmlns = parent.namespace_();
    rich_text t;

    while (in_element(parent))
    {
        auto text_element = expect_start_element(xml::content::mixed);
        const auto xml_space = qn("xml", "space");
        const auto preserve_space = parser().attribute_present(xml_space)
            ? parser().attribute(xml_space) == "preserve"
            : false;
        skip_attributes();
        auto text = read_text();

        if (text_element == xml::qname(xmlns, "t"))
        {
            t.plain_text(text, preserve_space);
        }
        else if (text_element == xml::qname(xmlns, "r"))
        {
            rich_text_run run;
            run.preserve_space = preserve_space;

            while (in_element(xml::qname(xmlns, "r")))
            {
                auto run_element = expect_start_element(xml::content::mixed);
                auto run_text = read_text();

                if (run_element == xml::qname(xmlns, "rPr"))
                {
                    run.second = xlnt::font();

                    while (in_element(xml::qname(xmlns, "rPr")))
                    {
                        auto current_run_property_element = expect_start_element(xml::content::simple);

                        if (current_run_property_element == xml::qname(xmlns, "sz"))
                        {
                            run.second.get().size(converter_.deserialise(parser().attribute("val")));
                        }
                        else if (current_run_property_element == xml::qname(xmlns, "rFont"))
                        {
                            run.second.get().name(parser().attribute("val"));
                        }
                        else if (current_run_property_element == xml::qname(xmlns, "color"))
                        {
                            run.second.get().color(read_color());
                        }
                        else if (current_run_property_element == xml::qname(xmlns, "family"))
                        {
                            run.second.get().family(parser().attribute<std::size_t>("val"));
                        }
                        else if (current_run_property_element == xml::qname(xmlns, "charset"))
                        {
                            run.second.get().charset(parser().attribute<std::size_t>("val"));
                        }
                        else if (current_run_property_element == xml::qname(xmlns, "scheme"))
                        {
                            run.second.get().scheme(parser().attribute("val"));
                        }
                        else if (current_run_property_element == xml::qname(xmlns, "b"))
                        {
                            run.second.get().bold(parser().attribute_present("val")
                                    ? is_true(parser().attribute("val"))
                                    : true);
                        }
                        else if (current_run_property_element == xml::qname(xmlns, "i"))
                        {
                            run.second.get().italic(parser().attribute_present("val")
                                    ? is_true(parser().attribute("val"))
                                    : true);
                        }
                        else if (current_run_property_element == xml::qname(xmlns, "u"))
                        {
                            if (parser().attribute_present("val"))
                            {
                                run.second.get().underline(parser().attribute<font::underline_style>("val"));
                            }
                            else
                            {
                                run.second.get().underline(font::underline_style::single);
                            }
                        }
                        else if (current_run_property_element == xml::qname(xmlns, "strike"))
                        {
                            run.second.get().strikethrough(parser().attribute_present("val")
                                    ? is_true(parser().attribute("val"))
                                    : true);
                        }
                        else
                        {
                            unexpected_element(current_run_property_element);
                        }

                        expect_end_element(current_run_property_element);
                        read_text();
                    }
                }
                else if (run_element == xml::qname(xmlns, "t"))
                {
                    run.first = run_text;
                }
                else
                {
                    unexpected_element(run_element);
                }

                read_text();
                expect_end_element(run_element);
                read_text();
            }

            t.add_run(run);
        }
        else if (text_element == xml::qname(xmlns, "rPh"))
        {
            phonetic_run pr;
            pr.start = parser().attribute<std::uint32_t>("sb");
            pr.end = parser().attribute<std::uint32_t>("eb");

            expect_start_element(xml::qname(xmlns, "t"), xml::content::simple);
            pr.text = read_text();

            if (parser().attribute_present(xml_space))
            {
                pr.preserve_space = parser().attribute(xml_space) == "preserve";
            }

            expect_end_element(xml::qname(xmlns, "t"));

            t.add_phonetic_run(pr);
        }
        else if (text_element == xml::qname(xmlns, "phoneticPr"))
        {
            phonetic_pr ph(parser().attribute<phonetic_pr::font_id_t>("fontId"));
            if (parser().attribute_present("type"))
            {
                ph.type(phonetic_pr::type_from_string(parser().attribute("type")));
            }
            if (parser().attribute_present("alignment"))
            {
                ph.alignment(phonetic_pr::alignment_from_string(parser().attribute("alignment")));
            }
            t.phonetic_properties(ph);
        }
        else
        {
            unexpected_element(text_element);
        }

        read_text();
        expect_end_element(text_element);
    }

    return t;
}

xlnt::color xlsx_consumer::read_color()
{
    xlnt::color result;

    if (parser().attribute_present("auto") && is_true(parser().attribute("auto")))
    {
        result.auto_(true);
        return result;
    }

    if (parser().attribute_present("rgb"))
    {
        result = xlnt::rgb_color(parser().attribute("rgb"));
    }
    else if (parser().attribute_present("theme"))
    {
        result = xlnt::theme_color(parser().attribute<std::size_t>("theme"));
    }
    else if (parser().attribute_present("indexed"))
    {
        result = xlnt::indexed_color(parser().attribute<std::size_t>("indexed"));
    }

    if (parser().attribute_present("tint"))
    {
        result.tint(converter_.deserialise(parser().attribute("tint")));
    }

    return result;
}

manifest &xlsx_consumer::manifest()
{
    return target_.manifest();
}

} // namespace detail
} // namespace xlnt
