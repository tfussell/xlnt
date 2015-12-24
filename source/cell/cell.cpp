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
#include <sstream>

#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/cell/comment.hpp>
#include <xlnt/packaging/document_properties.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/serialization/encoding.hpp>
#include <xlnt/styles/color.hpp>
#include <xlnt/styles/style.hpp>
#include <xlnt/utils/date.hpp>
#include <xlnt/utils/datetime.hpp>
#include <xlnt/utils/time.hpp>
#include <xlnt/utils/timedelta.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/utils/utf8string.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/column_properties.hpp>
#include <xlnt/worksheet/row_properties.hpp>
#include <xlnt/worksheet/worksheet.hpp>

#include <detail/cell_impl.hpp>
#include <detail/comment_impl.hpp>

namespace xlnt {

const std::unordered_map<std::string, int> &cell::error_codes()
{
    static const std::unordered_map<std::string, int> *codes =
    new std::unordered_map<std::string, int>({ { "#NULL!", 0 }, { "#DIV/0!", 1 }, { "#VALUE!", 2 },
                                                                { "#REF!", 3 },  { "#NAME?", 4 },  { "#NUM!", 5 },
                                                                { "#N/A!", 6 } });

    return *codes;
};

std::string cell::check_string(const std::string &to_check)
{
    // so we can modify it
    std::string s = to_check;

    if (s.size() == 0)
    {
        return s;
    }

    auto wb_encoding = get_parent().get_parent().get_encoding();

    //XXX: use utfcpp for this!
    switch(wb_encoding)
    {
    case encoding::latin1: break; // all bytes are valid in latin1
    case encoding::ascii:
        for (char c : s)
        {
            if (c < 0)
            {
                throw xlnt::unicode_decode_error(c);
            }
        }
        break;
    case encoding::utf8:
    {
	if(!utf8string::is_valid(s))
        {
            throw xlnt::unicode_decode_error('0');
        } 
        break;
    }
    case encoding::utf16:
    {
	if(!utf8string::from_utf16(s).is_valid())
        {
            throw xlnt::unicode_decode_error('0');
        } 
        break;
    }
    case encoding::utf32:
    {
	if(!utf8string::from_utf32(s).is_valid())
        {
            throw xlnt::unicode_decode_error('0');
        } 
        break;
    }
    default:
        // other encodings not supported yet
        break;
    } // switch(wb_encoding)

    // check encoding?
    if (s.size() > 32767)
    {
        s = s.substr(0, 32767); // max string length in Excel
    }

    for (char c : s)
    {
        if (c >= 0 && (c <= 8 || c == 11 || c == 12 || (c >= 14 && c <= 31)))
        {
            throw xlnt::illegal_character_error(c);
        }
    }

    return s;
}

cell::cell() : d_(nullptr)
{
}

cell::cell(detail::cell_impl *d) : d_(d)
{
}

cell::cell(worksheet worksheet, const cell_reference &reference) : d_(nullptr)
{
    cell self = worksheet.get_cell(reference);
    d_ = self.d_;
}

template <typename T>
cell::cell(worksheet worksheet, const cell_reference &reference, const T &initial_value) : cell(worksheet, reference)
{
    set_value(initial_value);
}

bool cell::garbage_collectible() const
{
    return !(get_data_type() != type::null || is_merged() || has_comment() || has_formula() || has_style());
}

template <>
XLNT_FUNCTION void cell::set_value(bool b)
{
    d_->value_numeric_ = b ? 1 : 0;
    d_->type_ = type::boolean;
}

template <>
XLNT_FUNCTION void cell::set_value(std::int8_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_FUNCTION void cell::set_value(std::int16_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_FUNCTION void cell::set_value(std::int32_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_FUNCTION void cell::set_value(std::int64_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_FUNCTION void cell::set_value(std::uint8_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_FUNCTION void cell::set_value(std::uint16_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_FUNCTION void cell::set_value(std::uint32_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_FUNCTION void cell::set_value(std::uint64_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

#ifdef _MSC_VER
template <>
XLNT_FUNCTION void cell::set_value(unsigned long i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}
#endif

#ifdef __linux
template <>
XLNT_FUNCTION void cell::set_value(long long i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_FUNCTION void cell::set_value(unsigned long long i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}
#endif

template <>
XLNT_FUNCTION void cell::set_value(float f)
{
    d_->value_numeric_ = static_cast<long double>(f);
    d_->type_ = type::numeric;
}

template <>
XLNT_FUNCTION void cell::set_value(double d)
{
    d_->value_numeric_ = static_cast<long double>(d);
    d_->type_ = type::numeric;
}

template <>
XLNT_FUNCTION void cell::set_value(long double d)
{
    d_->value_numeric_ = static_cast<long double>(d);
    d_->type_ = type::numeric;
}

template <>
XLNT_FUNCTION void cell::set_value(std::string s)
{
    d_->set_string(check_string(s), get_parent().get_parent().get_guess_types());

	if (get_data_type() == type::string && !s.empty())
	{
		get_parent().get_parent().add_shared_string(s);
	}
}

template <>
XLNT_FUNCTION void cell::set_value(char const *c)
{
    set_value(std::string(c));
}

template <>
XLNT_FUNCTION void cell::set_value(cell c)
{
    d_->type_ = c.d_->type_;
    d_->value_numeric_ = c.d_->value_numeric_;
    d_->value_string_ = c.d_->value_string_;
    d_->hyperlink_ = c.d_->hyperlink_;
    d_->has_hyperlink_ = c.d_->has_hyperlink_;
    d_->formula_ = c.d_->formula_;
    d_->style_id_ = c.d_->style_id_;
    set_comment(c.get_comment());
}

template <>
XLNT_FUNCTION void cell::set_value(date d)
{
    d_->type_ = type::numeric;
    d_->value_numeric_ = d.to_number(get_base_date());
    set_number_format(number_format::date_yyyymmdd2());
}

template <>
XLNT_FUNCTION void cell::set_value(datetime d)
{
    d_->type_ = type::numeric;
    d_->value_numeric_ = d.to_number(get_base_date());
    set_number_format(number_format::date_datetime());
}

template <>
XLNT_FUNCTION void cell::set_value(time t)
{
    d_->type_ = type::numeric;
    d_->value_numeric_ = t.to_number();
    set_number_format(number_format::date_time6());
}

template <>
XLNT_FUNCTION void cell::set_value(timedelta t)
{
    d_->type_ = type::numeric;
    d_->value_numeric_ = t.to_number();
    set_number_format(number_format::date_timedelta());
}

row_t cell::get_row() const
{
    return d_->row_;
}

column_t cell::get_column() const
{
    return d_->column_;
}

void cell::set_merged(bool merged)
{
    d_->is_merged_ = merged;
}

bool cell::is_merged() const
{
    return d_->is_merged_;
}

bool cell::is_date() const
{
    return get_data_type() == type::numeric && get_number_format().is_date_format();
}

cell_reference cell::get_reference() const
{
    return { d_->column_, d_->row_ };
}

bool cell::operator==(std::nullptr_t) const
{
    return d_ == nullptr;
}

bool cell::operator==(const cell &comparand) const
{
    return d_ == comparand.d_;
}

cell &cell::operator=(const cell &rhs)
{
    *d_ = *rhs.d_;
    return *this;
}

bool operator<(cell left, cell right)
{
    return left.get_reference() < right.get_reference();
}

std::string cell::to_repr() const
{
    return "<Cell " + worksheet(d_->parent_).get_title() + "." + get_reference().to_string() + ">";
}

relationship cell::get_hyperlink() const
{
    if (!d_->has_hyperlink_)
    {
        throw std::runtime_error("no hyperlink set");
    }

    return d_->hyperlink_;
}

bool cell::has_hyperlink() const
{
    return d_->has_hyperlink_;
}

void cell::set_hyperlink(const std::string &hyperlink)
{
    if (hyperlink.length() == 0 || std::find(hyperlink.begin(), hyperlink.end(), ':') == hyperlink.end())
    {
        throw data_type_exception();
    }

    d_->has_hyperlink_ = true;
    d_->hyperlink_ = worksheet(d_->parent_).create_relationship(relationship::type::hyperlink, hyperlink);

    if (get_data_type() == type::null)
    {
        set_value(hyperlink);
    }
}

void cell::set_formula(const std::string &formula)
{
    if (formula.length() == 0)
    {
        throw data_type_exception();
    }

    if (formula[0] == '=')
    {
        d_->formula_ = formula.substr(1);
    }
    else
    {
        d_->formula_ = formula;
    }
}

bool cell::has_formula() const
{
    return !d_->formula_.empty();
}

std::string cell::get_formula() const
{
    if (d_->formula_.empty())
    {
        throw data_type_exception();
    }

    return d_->formula_;
}

void cell::clear_formula()
{
    d_->formula_.clear();
}

void cell::set_comment(const xlnt::comment &c)
{
    if (c.d_ != d_->comment_.get())
    {
        throw xlnt::attribute_error();
    }

    if (!has_comment())
    {
        get_parent().increment_comments();
    }

    *get_comment().d_ = *c.d_;
}

void cell::clear_comment()
{
    if (has_comment())
    {
        get_parent().decrement_comments();
    }

    d_->comment_ = nullptr;
}

bool cell::has_comment() const
{
    return d_->comment_ != nullptr;
}

void cell::set_error(const std::string &error)
{
    if (error.length() == 0 || error[0] != '#')
    {
        throw data_type_exception();
    }

    d_->value_string_ = error;
    d_->type_ = type::error;
}

cell cell::offset(int column, int row)
{
    return get_parent().get_cell(cell_reference(d_->column_ + column, d_->row_ + row));
}

worksheet cell::get_parent()
{
    return worksheet(d_->parent_);
}

const worksheet cell::get_parent() const
{
    return worksheet(d_->parent_);
}

comment cell::get_comment()
{
    if (d_->comment_ == nullptr)
    {
        d_->comment_.reset(new detail::comment_impl());
        get_parent().increment_comments();
    }

    return comment(d_->comment_.get());
}

//TODO: this shares a lot of code with worksheet::get_point_pos, try to reduce repition
std::pair<int, int> cell::get_anchor() const
{
    static const double DefaultColumnWidth = 51.85;
    static const double DefaultRowHeight = 15.0;

    auto points_to_pixels = [](long double value, long double dpi)
    {
        return static_cast<int>(std::ceil(value * dpi / 72));
    };
    
    auto left_columns = d_->column_ - 1;
    int left_anchor = 0;
    auto default_width = points_to_pixels(DefaultColumnWidth, 96.0);

    for (column_t column_index = 1; column_index <= left_columns; column_index++)
    {
        if (get_parent().has_column_properties(column_index))
        {
            auto cdw = get_parent().get_column_properties(column_index).width;

            if (cdw > 0)
            {
                left_anchor += points_to_pixels(cdw, 96.0);
                continue;
            }
        }

        left_anchor += default_width;
    }

    auto top_rows = d_->row_ - 1;
    int top_anchor = 0;
    auto default_height = points_to_pixels(DefaultRowHeight, 96.0);

    for (row_t row_index = 1; row_index <= top_rows; row_index++)
    {
        if (get_parent().has_row_properties(row_index))
        {
            auto rdh = get_parent().get_row_properties(row_index).height;

            if (rdh > 0)
            {
                top_anchor += points_to_pixels(rdh, 96.0);
                continue;
            }
        }

        top_anchor += default_height;
    }

    return { left_anchor, top_anchor };
}

cell::type cell::get_data_type() const
{
    return d_->type_;
}

void cell::set_data_type(type t)
{
    d_->type_ = t;
}

const number_format &cell::get_number_format() const
{
    if (d_->has_style_)
    {
        return get_parent().get_parent().get_number_format(d_->style_id_);
    }
    else
    {
        return get_parent().get_parent().get_number_formats().front();
    }
}

const font &cell::get_font() const
{
    if (d_->has_style_)
    {
        auto font_id = get_parent().get_parent().get_styles()[d_->style_id_].get_font_id();
        return get_parent().get_parent().get_font(font_id);
    }
    else
    {
        return get_parent().get_parent().get_font(0);
    }
}

const fill &cell::get_fill() const
{
    return get_parent().get_parent().get_fill(d_->style_id_);
}

const border &cell::get_border() const
{
    return get_parent().get_parent().get_border(d_->style_id_);
}

const alignment &cell::get_alignment() const
{
    return get_parent().get_parent().get_alignment(d_->style_id_);
}

const protection &cell::get_protection() const
{
    return get_parent().get_parent().get_protection(d_->style_id_);
}

bool cell::pivot_button() const
{
    return get_parent().get_parent().get_pivot_button(d_->style_id_);
}

bool cell::quote_prefix() const
{
    return get_parent().get_parent().get_quote_prefix(d_->style_id_);
}

void cell::clear_value()
{
    d_->value_numeric_ = 0;
    d_->value_string_.clear();
    d_->formula_.clear();
    d_->type_ = cell::type::null;
}

template <>
XLNT_FUNCTION bool cell::get_value() const
{
    return d_->value_numeric_ != 0;
}

template <>
XLNT_FUNCTION std::int8_t cell::get_value() const
{
    return static_cast<std::int8_t>(d_->value_numeric_);
}

template <>
XLNT_FUNCTION std::int16_t cell::get_value() const
{
    return static_cast<std::int16_t>(d_->value_numeric_);
}

template <>
XLNT_FUNCTION std::int32_t cell::get_value() const
{
    return static_cast<std::int32_t>(d_->value_numeric_);
}

template <>
XLNT_FUNCTION std::int64_t cell::get_value() const
{
    return static_cast<std::int64_t>(d_->value_numeric_);
}

template <>
XLNT_FUNCTION std::uint8_t cell::get_value() const
{
    return static_cast<std::uint8_t>(d_->value_numeric_);
}

template <>
XLNT_FUNCTION std::uint16_t cell::get_value() const
{
    return static_cast<std::uint16_t>(d_->value_numeric_);
}

template <>
XLNT_FUNCTION std::uint32_t cell::get_value() const
{
    return static_cast<std::uint32_t>(d_->value_numeric_);
}

template <>
XLNT_FUNCTION std::uint64_t cell::get_value() const
{
    return static_cast<std::uint64_t>(d_->value_numeric_);
}

#ifdef __linux
template <>
XLNT_FUNCTION long long int cell::get_value() const
{
    return static_cast<long long int>(d_->value_numeric_);
}
#endif

template <>
XLNT_FUNCTION float cell::get_value() const
{
    return static_cast<float>(d_->value_numeric_);
}

template <>
XLNT_FUNCTION double cell::get_value() const
{
    return static_cast<double>(d_->value_numeric_);
}

template <>
XLNT_FUNCTION long double cell::get_value() const
{
    return d_->value_numeric_;
}

template <>
XLNT_FUNCTION time cell::get_value() const
{
    return time::from_number(d_->value_numeric_);
}

template <>
XLNT_FUNCTION datetime cell::get_value() const
{
    return datetime::from_number(d_->value_numeric_, get_base_date());
}

template <>
XLNT_FUNCTION date cell::get_value() const
{
    return date::from_number(static_cast<int>(d_->value_numeric_), get_base_date());
}

template <>
XLNT_FUNCTION timedelta cell::get_value() const
{
    return timedelta::from_number(d_->value_numeric_);
}

void cell::set_font(const font &font_)
{
    d_->has_style_ = true;
    d_->style_id_ = get_parent().get_parent().set_font(font_, d_->style_id_);
}

void cell::set_number_format(const number_format &number_format_)
{
    d_->has_style_ = true;
    d_->style_id_ = get_parent().get_parent().set_number_format(number_format_, d_->style_id_);
}

template <>
XLNT_FUNCTION std::string cell::get_value() const
{
    return d_->value_string_;
}

bool cell::has_value() const
{
    return d_->type_ != cell::type::null;
}

std::string cell::to_string() const
{
    auto nf = get_number_format();

    switch (get_data_type())
    {
    case cell::type::null:
        return "";
    case cell::type::numeric:
        return get_number_format().format(get_value<long double>(), get_base_date());
    case cell::type::string:
    case cell::type::formula:
    case cell::type::error:
        return get_number_format().format(get_value<std::string>());
    case cell::type::boolean:
        return get_value<long double>() == 0 ? "FALSE" : "TRUE";
    default:
        return "";
    }
}

std::size_t cell::get_style_id() const
{
    return d_->style_id_;
}

bool cell::has_style() const
{
    return d_->has_style_;
}

void cell::set_style_id(std::size_t style_id)
{
    d_->style_id_ = style_id;
    d_->has_style_ = true;
}

calendar cell::get_base_date() const
{
    return get_parent().get_parent().get_properties().excel_base_date;
}

std::ostream &operator<<(std::ostream &stream, const xlnt::cell &cell)
{
    return stream << cell.to_string();
}

} // namespace xlnt
