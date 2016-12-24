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
#include <cmath>
#include <sstream>

#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/cell/comment.hpp>
#include <xlnt/cell/rich_text.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/color.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/format.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/styles/style.hpp>
#include <xlnt/utils/date.hpp>
#include <xlnt/utils/datetime.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/utils/time.hpp>
#include <xlnt/utils/timedelta.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/column_properties.hpp>
#include <xlnt/worksheet/row_properties.hpp>
#include <xlnt/worksheet/worksheet.hpp>

#include <detail/cell_impl.hpp>
#include <detail/stylesheet.hpp>

namespace {

std::pair<bool, long double> cast_numeric(const std::string &s)
{
    auto str_end = static_cast<char *>(nullptr);
    auto result = std::strtold(s.c_str(), &str_end);

    return (str_end != s.c_str() + s.size())
        ? std::make_pair(false, 0.0L)
        : std::make_pair(true, result);
}

std::pair<bool, long double> cast_percentage(const std::string &s)
{
    if (s.back() == '%')
    {
        auto number = cast_numeric(s.substr(0, s.size() - 1));

        if (number.first)
        {
            return {true, number.second / 100};
        }
    }

    return {false, 0};
}

std::pair<bool, xlnt::time> cast_time(const std::string &s)
{
    xlnt::time result;

    std::vector<std::string> time_components;
    std::size_t prev = 0;
    auto colon_index = s.find(':');

    while (colon_index != std::string::npos)
    {
        time_components.push_back(s.substr(prev, colon_index - prev));
        prev = colon_index + 1;
        colon_index = s.find(':', colon_index + 1);
    }

    time_components.push_back(s.substr(prev, colon_index - prev));

    if (time_components.size() < 2 || time_components.size() > 3)
    {
        return {false, result};
    }

    std::vector<double> numeric_components;

    for (auto component : time_components)
    {
        if (component.empty() || (component.substr(0, component.find('.')).size() > 2))
        {
            return {false, result};
        }

        for (auto d : component)
        {
            if (!(d >= '0' && d <= '9') && d != '.')
            {
                return {false, result};
            }
        }

        auto without_leading_zero = component.front() == '0' ? component.substr(1) : component;
        auto numeric = std::stod(without_leading_zero);

        numeric_components.push_back(numeric);
    }

    result.hour = static_cast<int>(numeric_components[0]);
    result.minute = static_cast<int>(numeric_components[1]);

    if (std::fabs(static_cast<double>(result.minute) - numeric_components[1]) > std::numeric_limits<double>::epsilon())
    {
        result.minute = result.hour;
        result.hour = 0;
        result.second = static_cast<int>(numeric_components[1]);
        result.microsecond = static_cast<int>((numeric_components[1] - result.second) * 1E6);
    }
    else if (numeric_components.size() > 2)
    {
        result.second = static_cast<int>(numeric_components[2]);
        result.microsecond = static_cast<int>((numeric_components[2] - result.second) * 1E6);
    }

    return {true, result};
}

} // namespace

namespace xlnt {

const std::unordered_map<std::string, int> &cell::error_codes()
{
    static const auto *codes = new std::unordered_map<std::string, int>{
        {"#NULL!", 0},
        {"#DIV/0!", 1},
        {"#VALUE!", 2},
        {"#REF!", 3},
        {"#NAME?", 4},
        {"#NUM!", 5},
        {"#N/A!", 6}};

    return *codes;
}

std::string cell::check_string(const std::string &to_check)
{
    // so we can modify it
    std::string s = to_check;

    if (s.size() == 0)
    {
        return s;
    }
    else if (s.size() > 32767)
    {
        s = s.substr(0, 32767); // max string length in Excel
    }

    for (char c : s)
    {
        if (c >= 0 && (c <= 8 || c == 11 || c == 12 || (c >= 14 && c <= 31)))
        {
            throw illegal_character(c);
        }
    }

    return s;
}

cell::cell(detail::cell_impl *d)
    : d_(d)
{
}

bool cell::garbage_collectible() const
{
    return !(data_type() != type::null || is_merged() || has_formula() || has_format());
}

template <>
XLNT_API void cell::value(std::nullptr_t)
{
    d_->type_ = type::null;
}

template <>
XLNT_API void cell::value(bool b)
{
    d_->value_numeric_ = b ? 1 : 0;
    d_->type_ = type::boolean;
}

template <>
XLNT_API void cell::value(std::int8_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::value(std::int16_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::value(std::int32_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::value(std::int64_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::value(std::uint8_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::value(std::uint16_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::value(std::uint32_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::value(std::uint64_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

#ifdef _MSC_VER
template <>
XLNT_API void cell::value(unsigned long i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}
#endif

#ifdef __linux
template <>
XLNT_API void cell::value(long long i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::value(unsigned long long i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}
#endif

template <>
XLNT_API void cell::value(float f)
{
    d_->value_numeric_ = static_cast<long double>(f);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::value(double d)
{
    d_->value_numeric_ = static_cast<long double>(d);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::value(long double d)
{
    d_->value_numeric_ = static_cast<long double>(d);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::value(std::string s)
{
    s = check_string(s);

    if (s.size() > 1 && s.front() == '=')
    {
        d_->type_ = type::formula;
        formula(s);
    }
    else if (cell::error_codes().find(s) != cell::error_codes().end())
    {
        error(s);
    }
    else
    {
        d_->type_ = type::string;
        d_->value_text_.plain_text(s);

        if (s.size() > 0)
        {
            workbook().add_shared_string(d_->value_text_);
        }
    }
}

template <>
XLNT_API void cell::value(rich_text text)
{
    if (text.runs().size() == 1 && !text.runs().front().second.is_set())
    {
        value(text.plain_text());
    }
    else
    {
        d_->type_ = type::string;
        d_->value_text_ = text;
        workbook().add_shared_string(text);
    }
}

template <>
XLNT_API void cell::value(char const *c)
{
    value(std::string(c));
}

template <>
XLNT_API void cell::value(cell c)
{
    d_->type_ = c.d_->type_;
    d_->value_numeric_ = c.d_->value_numeric_;
    d_->value_text_ = c.d_->value_text_;
    d_->hyperlink_ = c.d_->hyperlink_;
    d_->formula_ = c.d_->formula_;
    d_->format_ = c.d_->format_;
}

template <>
XLNT_API void cell::value(date d)
{
    d_->type_ = type::numeric;
    d_->value_numeric_ = d.to_number(base_date());
    number_format(number_format::date_yyyymmdd2());
}

template <>
XLNT_API void cell::value(datetime d)
{
    d_->type_ = type::numeric;
    d_->value_numeric_ = d.to_number(base_date());
    number_format(number_format::date_datetime());
}

template <>
XLNT_API void cell::value(time t)
{
    d_->type_ = type::numeric;
    d_->value_numeric_ = t.to_number();
    number_format(number_format::date_time6());
}

template <>
XLNT_API void cell::value(timedelta t)
{
    d_->type_ = type::numeric;
    d_->value_numeric_ = t.to_number();
    number_format(xlnt::number_format("[hh]:mm:ss"));
}

row_t cell::row() const
{
    return d_->row_;
}

column_t cell::column() const
{
    return d_->column_;
}

void cell::merged(bool merged)
{
    d_->is_merged_ = merged;
}

bool cell::is_merged() const
{
    return d_->is_merged_;
}

bool cell::is_date() const
{
    return data_type() == type::numeric && has_format() && number_format().is_date_format();
}

cell_reference cell::reference() const
{
    return {d_->column_, d_->row_};
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
    d_->column_ = rhs.d_->column_;
    d_->format_ = rhs.d_->format_;
    d_->formula_ = rhs.d_->formula_;
    d_->hyperlink_ = rhs.d_->hyperlink_;
    d_->is_merged_ = rhs.d_->is_merged_;
    d_->parent_ = rhs.d_->parent_;
    d_->row_ = rhs.d_->row_;
    d_->type_ = rhs.d_->type_;
    d_->value_numeric_ = rhs.d_->value_numeric_;
    d_->value_text_ = rhs.d_->value_text_;

    return *this;
}

std::string cell::hyperlink() const
{
    return d_->hyperlink_.get();
}

void cell::hyperlink(const std::string &hyperlink)
{
    if (hyperlink.length() == 0 || std::find(hyperlink.begin(), hyperlink.end(), ':') == hyperlink.end())
    {
        throw invalid_parameter();
    }

    d_->hyperlink_ = hyperlink;

    if (data_type() == type::null)
    {
        value(hyperlink);
    }
}

void cell::formula(const std::string &formula)
{
    if (formula.empty())
    {
        throw invalid_parameter();
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
    return d_->formula_.is_set();
}

std::string cell::formula() const
{
    return d_->formula_.get();
}

void cell::clear_formula()
{
    d_->formula_.clear();
}

void cell::error(const std::string &error)
{
    if (error.length() == 0 || error[0] != '#')
    {
        throw invalid_data_type();
    }

    d_->value_text_.plain_text(error);
    d_->type_ = type::error;
}

cell cell::offset(int column, int row)
{
    return worksheet().cell(reference().make_offset(column, row));
}

worksheet cell::worksheet()
{
    return xlnt::worksheet(d_->parent_);
}

const worksheet cell::worksheet() const
{
    return xlnt::worksheet(d_->parent_);
}

workbook &cell::workbook()
{
    return worksheet().workbook();
}

const workbook &cell::workbook() const
{
    return worksheet().workbook();
}

std::pair<int, int> cell::anchor() const
{
    int left = 0;

    for (column_t column_index = 1; column_index <= d_->column_ - 1; column_index++)
    {
        left += worksheet().cell(column_index, row()).width();
    }

    int top = 0;

    for (row_t row_index = 1; row_index <= d_->row_ - 1; row_index++)
    {
        top += worksheet().cell(column(), row_index).height();
    }

    return {left, top};
}

cell::type cell::data_type() const
{
    return d_->type_;
}

void cell::data_type(type t)
{
    d_->type_ = t;
}

number_format cell::computed_number_format() const
{
    return xlnt::number_format();
}

font cell::computed_font() const
{
    return xlnt::font();
}

fill cell::computed_fill() const
{
    return xlnt::fill();
}

border cell::computed_border() const
{
    return xlnt::border();
}

alignment cell::computed_alignment() const
{
    return xlnt::alignment();
}

protection cell::computed_protection() const
{
    return xlnt::protection();
}

void cell::clear_value()
{
    d_->value_numeric_ = 0;
    d_->value_text_.clear();
    d_->formula_.clear();
    d_->type_ = cell::type::null;
}

template <>
XLNT_API bool cell::value() const
{
    return d_->value_numeric_ != 0.L;
}

template <>
XLNT_API std::int8_t cell::value() const
{
    return static_cast<std::int8_t>(d_->value_numeric_);
}

template <>
XLNT_API std::int16_t cell::value() const
{
    return static_cast<std::int16_t>(d_->value_numeric_);
}

template <>
XLNT_API std::int32_t cell::value() const
{
    return static_cast<std::int32_t>(d_->value_numeric_);
}

template <>
XLNT_API std::int64_t cell::value() const
{
    return static_cast<std::int64_t>(d_->value_numeric_);
}

template <>
XLNT_API std::uint8_t cell::value() const
{
    return static_cast<std::uint8_t>(d_->value_numeric_);
}

template <>
XLNT_API std::uint16_t cell::value() const
{
    return static_cast<std::uint16_t>(d_->value_numeric_);
}

template <>
XLNT_API std::uint32_t cell::value() const
{
    return static_cast<std::uint32_t>(d_->value_numeric_);
}

template <>
XLNT_API std::uint64_t cell::value() const
{
    return static_cast<std::uint64_t>(d_->value_numeric_);
}

#ifdef __linux
template <>
XLNT_API long long cell::value() const
{
    return static_cast<long long>(d_->value_numeric_);
}

template <>
XLNT_API unsigned long long cell::value() const
{
    return static_cast<unsigned long long>(d_->value_numeric_);
}
#endif

template <>
XLNT_API float cell::value() const
{
    return static_cast<float>(d_->value_numeric_);
}

template <>
XLNT_API double cell::value() const
{
    return static_cast<double>(d_->value_numeric_);
}

template <>
XLNT_API long double cell::value() const
{
    return d_->value_numeric_;
}

template <>
XLNT_API time cell::value() const
{
    return time::from_number(d_->value_numeric_);
}

template <>
XLNT_API datetime cell::value() const
{
    return datetime::from_number(d_->value_numeric_, base_date());
}

template <>
XLNT_API date cell::value() const
{
    return date::from_number(static_cast<int>(d_->value_numeric_), base_date());
}

template <>
XLNT_API timedelta cell::value() const
{
    return timedelta::from_number(d_->value_numeric_);
}

void cell::alignment(const class alignment &alignment_)
{
    auto new_format = has_format() ? modifiable_format() : workbook().create_format();
    format(new_format.alignment(alignment_, true));
}

void cell::border(const class border &border_)
{
    auto new_format = has_format() ? modifiable_format() : workbook().create_format();
    format(new_format.border(border_, true));
}

void cell::fill(const class fill &fill_)
{
    auto new_format = has_format() ? modifiable_format() : workbook().create_format();
    format(new_format.fill(fill_, true));
}

void cell::font(const class font &font_)
{
    auto new_format = has_format() ? modifiable_format() : workbook().create_format();
    format(new_format.font(font_, true));
}

void cell::number_format(const class number_format &number_format_)
{
    auto new_format = has_format() ? modifiable_format() : workbook().create_format();
    format(new_format.number_format(number_format_, true));
}

void cell::protection(const class protection &protection_)
{
    auto new_format = has_format() ? modifiable_format() : workbook().create_format();
    format(new_format.protection(protection_, true));
}

template <>
XLNT_API std::string cell::value() const
{
    return d_->value_text_.plain_text();
}

template <>
XLNT_API rich_text cell::value() const
{
    return d_->value_text_;
}

bool cell::has_value() const
{
    return d_->type_ != cell::type::null;
}

std::string cell::to_string() const
{
    auto nf = computed_number_format();

    switch (data_type())
    {
    case cell::type::null:
        return "";
    case cell::type::numeric:
        return nf.format(value<long double>(), base_date());
    case cell::type::string:
    case cell::type::formula:
    case cell::type::error:
        return nf.format(value<std::string>());
    case cell::type::boolean:
        return value<long double>() == 0.L ? "FALSE" : "TRUE";
    }

    return "";
}

bool cell::has_format() const
{
    return d_->format_.is_set();
}

void cell::format(const class format new_format)
{
    if (has_format())
    {
        format().d_->references -= format().d_->references > 0 ? 1 : 0;
    }

    ++new_format.d_->references;
    d_->format_ = new_format.d_;
}

calendar cell::base_date() const
{
    return workbook().base_date();
}

XLNT_API std::ostream &operator<<(std::ostream &stream, const xlnt::cell &cell)
{
    return stream << cell.to_string();
}

void cell::value(const std::string &value_string, bool infer_type)
{
    value(value_string);

    if (!infer_type)
    {
        return;
    }

    auto percentage = cast_percentage(value_string);

    if (percentage.first)
    {
        d_->value_numeric_ = percentage.second;
        d_->type_ = cell::type::numeric;
        number_format(xlnt::number_format::percentage());
    }
    else
    {
        auto time = cast_time(value_string);

        if (time.first)
        {
            d_->type_ = cell::type::numeric;
            number_format(number_format::date_time6());
            d_->value_numeric_ = time.second.to_number();
        }
        else
        {
            auto numeric = cast_numeric(value_string);

            if (numeric.first)
            {
                d_->value_numeric_ = numeric.second;
                d_->type_ = cell::type::numeric;
            }
        }
    }
}

void cell::clear_format()
{
    format().d_->references -= format().d_->references > 0 ? 1 : 0;
    d_->format_.clear();
}

void cell::clear_style()
{
    if (has_format())
    {
        modifiable_format().clear_style();
    }
}

void cell::style(const class style &new_style)
{
    auto new_format = has_format() ? format() : workbook().create_format();
    format(new_format.style(new_style));
}

void cell::style(const std::string &style_name)
{
    style(workbook().style(style_name));
}

const style cell::style() const
{
    if (!has_format() || !format().has_style())
    {
        throw invalid_attribute();
    }

    return format().style();
}

bool cell::has_style() const
{
    return has_format() && format().has_style();
}

format cell::modifiable_format()
{
    if (!d_->format_)
    {
        throw invalid_attribute();
    }

    return xlnt::format(*d_->format_);
}

const format cell::format() const
{
    if (!d_->format_)
    {
        throw invalid_attribute();
    }

    return xlnt::format(*d_->format_);
}

alignment cell::alignment() const
{
    return format().alignment();
}

border cell::border() const
{
    return format().border();
}

fill cell::fill() const
{
    return format().fill();
}

font cell::font() const
{
    return format().font();
}

number_format cell::number_format() const
{
    return format().number_format();
}

protection cell::protection() const
{
    return format().protection();
}

bool cell::has_hyperlink() const
{
    return d_->hyperlink_;
}

// comment

bool cell::has_comment()
{
    return d_->comment_.is_set();
}

void cell::clear_comment()
{
    d_->comment_.clear();
}

class comment cell::comment()
{
    if (!has_comment())
    {
        throw xlnt::exception("cell has no comment");
    }

    return d_->comment_.get();
}

void cell::comment(const std::string &text, const std::string &author)
{
    comment(xlnt::comment(text, author));
}

void cell::comment(const std::string &text, const class font &comment_font, const std::string &author)
{
    xlnt::rich_text rich_comment_text(text, comment_font);
    comment(xlnt::comment(rich_comment_text, author));
}

void cell::comment(const class comment &new_comment)
{
    d_->comment_.set(new_comment);

    // offset comment 5 pixels down and 5 pixels right of the top right corner of the cell
    auto cell_position = anchor();
    cell_position.first += width() + 5;
    cell_position.second += 5;

    d_->comment_.get().position(cell_position.first, cell_position.second);
    d_->comment_.get().size(200, 100);

    worksheet().register_comments_in_manifest();
}

double cell::width() const
{
    return worksheet().column_width(column());
}

double cell::height() const
{
    return worksheet().row_height(row());
}

} // namespace xlnt
