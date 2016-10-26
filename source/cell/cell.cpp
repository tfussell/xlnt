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
#include <xlnt/cell/text.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/styles/color.hpp>
#include <xlnt/styles/format.hpp>
#include <xlnt/styles/style.hpp>
#include <xlnt/utils/date.hpp>
#include <xlnt/utils/datetime.hpp>
#include <xlnt/utils/time.hpp>
#include <xlnt/utils/timedelta.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/column_properties.hpp>
#include <xlnt/worksheet/row_properties.hpp>
#include <xlnt/worksheet/worksheet.hpp>

#include <detail/cell_impl.hpp>
#include <detail/comment_impl.hpp>


namespace {

std::pair<bool, long double> cast_numeric(const std::string &s)
{
	const char *str = s.c_str();
	char *str_end = nullptr;
	auto result = std::strtold(str, &str_end);
	if (str_end != str + s.size()) return{ false, 0 };
	return{ true, result };
}

std::pair<bool, long double> cast_percentage(const std::string &s)
{
	if (s.back() == '%')
	{
		auto number = cast_numeric(s.substr(0, s.size() - 1));

		if (number.first)
		{
			return{ true, number.second / 100 };
		}
	}

	return{ false, 0 };
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
		return{ false, result };
	}

	std::vector<double> numeric_components;

	for (auto component : time_components)
	{
		if (component.empty() || (component.substr(0, component.find('.')).size() > 2))
		{
			return{ false, result };
		}

		for (auto d : component)
		{
			if (!(d >= '0' && d <= '9') && d != '.')
			{
				return{ false, result };
			}
		}

		auto without_leading_zero = component.front() == '0' ? component.substr(1) : component;
		auto numeric = std::stod(without_leading_zero);

		numeric_components.push_back(numeric);
	}

	result.hour = static_cast<int>(numeric_components[0]);
	result.minute = static_cast<int>(numeric_components[1]);

	if (result.minute != numeric_components[1])
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

	return{ true, result };
}

} // namespace

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

cell::cell(detail::cell_impl *d) : d_(d)
{
}

bool cell::garbage_collectible() const
{
    return !(get_data_type() != type::null || is_merged() || has_formula() || has_format());
}

template <>
XLNT_API void cell::set_value(std::nullptr_t)
{
	d_->type_ = type::null;
}

template <>
XLNT_API void cell::set_value(bool b)
{
    d_->value_numeric_ = b ? 1 : 0;
    d_->type_ = type::boolean;
}

template <>
XLNT_API void cell::set_value(std::int8_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::set_value(std::int16_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::set_value(std::int32_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::set_value(std::int64_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::set_value(std::uint8_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::set_value(std::uint16_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::set_value(std::uint32_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::set_value(std::uint64_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

#ifdef _MSC_VER
template <>
XLNT_API void cell::set_value(unsigned long i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}
#endif

#ifdef __linux
template <>
XLNT_API void cell::set_value(long long i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::set_value(unsigned long long i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}
#endif

template <>
XLNT_API void cell::set_value(float f)
{
    d_->value_numeric_ = static_cast<long double>(f);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::set_value(double d)
{
    d_->value_numeric_ = static_cast<long double>(d);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::set_value(long double d)
{
    d_->value_numeric_ = static_cast<long double>(d);
    d_->type_ = type::numeric;
}

template <>
XLNT_API void cell::set_value(std::string s)
{
	s = check_string(s);

	if (s.size() > 1 && s.front() == '=')
	{
		d_->type_ = type::formula;
		set_formula(s);
	}
	else if (cell::error_codes().find(s) != cell::error_codes().end())
	{
		set_error(s);
	}
	else
	{
		d_->type_ = type::string;
        d_->value_text_.set_plain_string(s);
        
        if (s.size() > 0)
        {
            get_workbook().add_shared_string(d_->value_text_);
        }
	}

	if (get_workbook().get_guess_types())
	{
		guess_type_and_set_value(s);
	}
}

template <>
XLNT_API void cell::set_value(text t)
{
    if (t.get_runs().size() == 1 && !t.get_runs().front().has_formatting())
    {
        set_value(t.get_plain_string());
    }
    else
    {
        d_->type_ = type::string;
        d_->value_text_ = t;
        get_workbook().add_shared_string(t);
    }
}

template <>
XLNT_API void cell::set_value(char const *c)
{
    set_value(std::string(c));
}

template <>
XLNT_API void cell::set_value(cell c)
{
    d_->type_ = c.d_->type_;
    d_->value_numeric_ = c.d_->value_numeric_;
    d_->value_text_ = c.d_->value_text_;
    d_->hyperlink_ = c.d_->hyperlink_;
    d_->formula_ = c.d_->formula_;
    d_->format_id_ = c.d_->format_id_;
}

template <>
XLNT_API void cell::set_value(date d)
{
    d_->type_ = type::numeric;
    d_->value_numeric_ = d.to_number(get_base_date());
    set_number_format(number_format::date_yyyymmdd2());
}

template <>
XLNT_API void cell::set_value(datetime d)
{
    d_->type_ = type::numeric;
    d_->value_numeric_ = d.to_number(get_base_date());
    set_number_format(number_format::date_datetime());
}

template <>
XLNT_API void cell::set_value(time t)
{
    d_->type_ = type::numeric;
    d_->value_numeric_ = t.to_number();
    set_number_format(number_format::date_time6());
}

template <>
XLNT_API void cell::set_value(timedelta t)
{
    d_->type_ = type::numeric;
    d_->value_numeric_ = t.to_number();
    set_number_format(number_format("[hh]:mm:ss"));
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
    return get_data_type() == type::numeric 
		&& has_format()
		&& get_number_format().is_date_format();
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
	d_->column_ = rhs.d_->column_;
	d_->format_id_ = rhs.d_->format_id_;
	d_->formula_ = rhs.d_->formula_;
	d_->hyperlink_ = rhs.d_->hyperlink_;
	d_->is_merged_ = rhs.d_->is_merged_;
	d_->parent_ = rhs.d_->parent_;
	d_->row_ = rhs.d_->row_;
	d_->style_name_ = rhs.d_->style_name_;
	d_->type_ = rhs.d_->type_;
	d_->value_numeric_ = rhs.d_->value_numeric_;
	d_->value_text_ = rhs.d_->value_text_;

    return *this;
}

std::string cell::to_repr() const
{
    return "<Cell " + worksheet(d_->parent_).get_title() + "." + get_reference().to_string() + ">";
}

std::string cell::get_hyperlink() const
{
	return d_->hyperlink_.get();
}

void cell::set_hyperlink(const std::string &hyperlink)
{
    if (hyperlink.length() == 0 || std::find(hyperlink.begin(), hyperlink.end(), ':') == hyperlink.end())
    {
        throw invalid_parameter();
    }

	d_->hyperlink_ = hyperlink;

    if (get_data_type() == type::null)
    {
        set_value(hyperlink);
    }
}

void cell::set_formula(const std::string &formula)
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
	return d_->formula_;
}

std::string cell::get_formula() const
{
    return d_->formula_.get();
}

void cell::clear_formula()
{
    d_->formula_.clear();
}

void cell::set_error(const std::string &error)
{
    if (error.length() == 0 || error[0] != '#')
    {
        throw invalid_data_type();
    }

    d_->value_text_.set_plain_string(error);
    d_->type_ = type::error;
}

cell cell::offset(int column, int row)
{
    return get_worksheet().get_cell(cell_reference(d_->column_ + column, d_->row_ + row));
}

worksheet cell::get_worksheet()
{
    return worksheet(d_->parent_);
}

const worksheet cell::get_worksheet() const
{
    return worksheet(d_->parent_);
}

workbook &cell::get_workbook()
{
	return get_worksheet().get_workbook();
}

const workbook &cell::get_workbook() const
{
	return get_worksheet().get_workbook();
}

//TODO: this shares a lot of code with worksheet::get_point_pos, try to reduce repetion
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
        if (get_worksheet().has_column_properties(column_index))
        {
            auto cdw = get_worksheet().get_column_properties(column_index).width;

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
        if (get_worksheet().has_row_properties(row_index))
        {
            auto rdh = get_worksheet().get_row_properties(row_index).height;

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

number_format cell::get_computed_number_format() const
{
	return get_computed_format().number_format();
}

font cell::get_computed_font() const
{
	return get_computed_format().font();
}

fill cell::get_computed_fill() const
{
	return get_computed_format().fill();
}

border cell::get_computed_border() const
{
	return get_computed_format().border();
}

alignment cell::get_computed_alignment() const
{
	return get_computed_format().alignment();
}

protection cell::get_computed_protection() const
{
	return get_computed_format().protection();
}

void cell::clear_value()
{
    d_->value_numeric_ = 0;
    d_->value_text_.clear();
    d_->formula_.clear();
    d_->type_ = cell::type::null;
}

template <>
XLNT_API bool cell::get_value() const
{
    return d_->value_numeric_ != 0;
}

template <>
XLNT_API std::int8_t cell::get_value() const
{
    return static_cast<std::int8_t>(d_->value_numeric_);
}

template <>
XLNT_API std::int16_t cell::get_value() const
{
    return static_cast<std::int16_t>(d_->value_numeric_);
}

template <>
XLNT_API std::int32_t cell::get_value() const
{
    return static_cast<std::int32_t>(d_->value_numeric_);
}

template <>
XLNT_API std::int64_t cell::get_value() const
{
    return static_cast<std::int64_t>(d_->value_numeric_);
}

template <>
XLNT_API std::uint8_t cell::get_value() const
{
    return static_cast<std::uint8_t>(d_->value_numeric_);
}

template <>
XLNT_API std::uint16_t cell::get_value() const
{
    return static_cast<std::uint16_t>(d_->value_numeric_);
}

template <>
XLNT_API std::uint32_t cell::get_value() const
{
    return static_cast<std::uint32_t>(d_->value_numeric_);
}

template <>
XLNT_API std::uint64_t cell::get_value() const
{
    return static_cast<std::uint64_t>(d_->value_numeric_);
}

#ifdef __linux
template <>
XLNT_API long long cell::get_value() const
{
    return static_cast<long long>(d_->value_numeric_);
}

template <>
XLNT_API unsigned long long cell::get_value() const
{
    return static_cast<unsigned long long>(d_->value_numeric_);
}
#endif

template <>
XLNT_API float cell::get_value() const
{
    return static_cast<float>(d_->value_numeric_);
}

template <>
XLNT_API double cell::get_value() const
{
    return static_cast<double>(d_->value_numeric_);
}

template <>
XLNT_API long double cell::get_value() const
{
    return d_->value_numeric_;
}

template <>
XLNT_API time cell::get_value() const
{
    return time::from_number(d_->value_numeric_);
}

template <>
XLNT_API datetime cell::get_value() const
{
    return datetime::from_number(d_->value_numeric_, get_base_date());
}

template <>
XLNT_API date cell::get_value() const
{
    return date::from_number(static_cast<int>(d_->value_numeric_), get_base_date());
}

template <>
XLNT_API timedelta cell::get_value() const
{
    return timedelta::from_number(d_->value_numeric_);
}

void cell::set_border(const xlnt::border &border_)
{
	auto &format = get_workbook().create_format();
	format.border(border_, true);
	d_->format_id_ = format.id();
}

void cell::set_fill(const xlnt::fill &fill_)
{
	auto &format = get_workbook().create_format();
	format.fill(fill_, true);
	d_->format_id_ = format.id();
}

void cell::set_font(const font &font_)
{
	auto &format = get_workbook().create_format();
	format.font(font_, true);
	d_->format_id_ = format.id();
}

void cell::set_number_format(const number_format &number_format_)
{
	auto &format = get_workbook().create_format();
	format.number_format(number_format_, true);
	d_->format_id_ = format.id();
}

void cell::set_alignment(const xlnt::alignment &alignment_)
{
	auto &format = get_workbook().create_format();
	format.alignment(alignment_, true);
	d_->format_id_ = format.id();
}

void cell::set_protection(const xlnt::protection &protection_)
{
	auto &format = get_workbook().create_format();
	format.protection(protection_, true);
	d_->format_id_ = format.id();
}

template <>
XLNT_API std::string cell::get_value() const
{
    return d_->value_text_.get_plain_string();
}

template <>
XLNT_API text cell::get_value() const
{
    return d_->value_text_;
}

bool cell::has_value() const
{
    return d_->type_ != cell::type::null;
}

std::string cell::to_string() const
{
	auto nf = get_computed_number_format();

    switch (get_data_type())
    {
    case cell::type::null:
        return "";
    case cell::type::numeric:
        return nf.format(get_value<long double>(), get_base_date());
    case cell::type::string:
    case cell::type::formula:
    case cell::type::error:
        return nf.format(get_value<std::string>());
    case cell::type::boolean:
        return get_value<long double>() == 0 ? "FALSE" : "TRUE";
    default:
        return "";
    }
}

bool cell::has_format() const
{
	return d_->format_id_.is_set();
}

void cell::set_format(const format &new_format)
{
	d_->format_id_ = get_workbook().create_format().id();
}

calendar cell::get_base_date() const
{
	return get_workbook().get_base_date();
}

std::ostream &operator<<(std::ostream &stream, const xlnt::cell &cell)
{
    return stream << cell.to_string();
}

void cell::guess_type_and_set_value(const std::string &value)
{
	auto percentage = cast_percentage(value);

	if (percentage.first)
	{
		d_->value_numeric_ = percentage.second;
		d_->type_ = cell::type::numeric;
		set_number_format(xlnt::number_format::percentage());
	}
	else
	{
		auto time = cast_time(value);

		if (time.first)
		{
			d_->type_ = cell::type::numeric;
			set_number_format(number_format::date_time6());
			d_->value_numeric_ = time.second.to_number();
		}
		else
		{
			auto numeric = cast_numeric(value);

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
	d_->format_id_.clear();
}

void cell::clear_style()
{
	d_->style_name_.clear();
}

void cell::set_style(const style &new_style)
{
	d_->style_name_ = new_style.name();
}

void cell::set_style(const std::string &style_name)
{
	d_->style_name_ = get_workbook().get_style(style_name).name();
}

style cell::get_style() const
{
    if (!d_->style_name_)
    {
		throw invalid_attribute();
    }

    return get_workbook().get_style(*d_->style_name_);
}

bool cell::has_style() const
{
    return d_->style_name_;
}

base_format cell::get_computed_format() const
{
	base_format result;

	// Check style first

	if (has_style())
	{
		style cell_style = get_style();

		if (cell_style.alignment_applied()) result.alignment(cell_style.alignment(), true);
		if (cell_style.border_applied()) result.border(cell_style.border(), true);
		if (cell_style.fill_applied()) result.fill(cell_style.fill(), true);
		if (cell_style.font_applied()) result.font(cell_style.font(), true);
		if (cell_style.number_format_applied()) result.number_format(cell_style.number_format(), true);
		if (cell_style.protection_applied()) result.protection(cell_style.protection(), true);
	}

	// Cell format overrides style

	if (has_format())
	{
		format cell_format = get_format();

		if (cell_format.alignment_applied()) result.alignment(cell_format.alignment(), true);
		if (cell_format.border_applied()) result.border(cell_format.border(), true);
		if (cell_format.fill_applied()) result.fill(cell_format.fill(), true);
		if (cell_format.font_applied()) result.font(cell_format.font(), true);
		if (cell_format.number_format_applied()) result.number_format(cell_format.number_format(), true);
		if (cell_format.protection_applied()) result.protection(cell_format.protection(), true);
	}

	// Use defaults for any remaining non-applied components

	if (!result.alignment_applied()) result.alignment(alignment(), true);
	if (!result.border_applied()) result.border(border(), true);
	if (!result.fill_applied()) result.fill(fill(), true);
	if (!result.font_applied()) result.font(font(), true);
	if (!result.number_format_applied()) result.number_format(number_format(), true);
	if (!result.protection_applied()) result.protection(protection(), true);

	return result;
}

format &cell::get_format_internal()
{
	if (!d_->format_id_)
	{
		throw invalid_attribute();
	}

	return get_workbook().get_format(*d_->format_id_);
}

format cell::get_format() const
{
	if (!d_->format_id_)
	{
		throw invalid_attribute();
	}

	return get_workbook().get_format(*d_->format_id_);
}

alignment cell::get_alignment() const
{
	return get_format().alignment();
}

border cell::get_border() const
{
	return get_format().border();
}

fill cell::get_fill() const
{
	return get_format().fill();
}

font cell::get_font() const
{
	return get_format().font();
}

number_format cell::get_number_format() const
{
	return get_format().number_format();
}

protection cell::get_protection() const
{
	return get_format().protection();
}

bool cell::has_hyperlink() const
{
	return d_->hyperlink_;
}

} // namespace xlnt
