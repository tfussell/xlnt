#include <algorithm>
#include <locale>
#include <sstream>

#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/cell/comment.hpp>
#include <xlnt/common/datetime.hpp>
#include <xlnt/common/relationship.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/common/exceptions.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/workbook/document_properties.hpp>
#include <xlnt/common/exceptions.hpp>

#include "detail/cell_impl.hpp"
#include "detail/comment_impl.hpp"

namespace xlnt {
    
const xlnt::color xlnt::color::black(0);
const xlnt::color xlnt::color::white(1);
    
const std::unordered_map<std::string, int> cell::ErrorCodes =
{
    {"#NULL!", 0},
    {"#DIV/0!", 1},
    {"#VALUE!", 2},
    {"#REF!", 3},
    {"#NAME?", 4},
    {"#NUM!", 5},
    {"#N/A!", 6}
};

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
  
template<typename T>
cell::cell(worksheet worksheet, const cell_reference &reference, const T &initial_value) : cell(worksheet, reference)
{
    set_value(initial_value);
}
    
bool cell::garbage_collectible() const
{
    return !(get_data_type() != type::null || is_merged() || has_comment() || has_formula());
}


template<>
void cell::set_value(bool b)
{
    d_->is_date_ = false;
    d_->value_numeric_ = b ? 1 : 0;
    d_->type_ = type::boolean;
}

template<>
void cell::set_value(std::int8_t i)
{
    d_->is_date_ = false;
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template<>
void cell::set_value(std::int16_t i)
{
    d_->is_date_ = false;
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template<>
void cell::set_value(std::int32_t i)
{
    d_->is_date_ = false;
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template<>
void cell::set_value(std::int64_t i)
{
    d_->is_date_ = false;
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template<>
void cell::set_value(std::uint8_t i)
{
    d_->is_date_ = false;
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template<>
void cell::set_value(std::uint16_t i)
{
    d_->is_date_ = false;
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template<>
void cell::set_value(std::uint32_t i)
{
    d_->is_date_ = false;
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template<>
void cell::set_value(std::uint64_t i)
{
    d_->is_date_ = false;
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

#ifdef _WIN32
void cell::set_value(unsigned long i)
{
    d_->is_date_ = false;
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}
#endif

#ifdef __linux__
void cell::set_value(long long i)
{
    d_->is_date_ = false;
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

void cell::set_value(unsigned long long i)
{
    d_->is_date_ = false;
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}
#endif

template<>
void cell::set_value(float f)
{
    d_->is_date_ = false;
    d_->value_numeric_ = static_cast<long double>(f);
    d_->type_ = type::numeric;
}

template<>
void cell::set_value(double d)
{
    d_->is_date_ = false;
    d_->value_numeric_ = static_cast<long double>(d);
    d_->type_ = type::numeric;
}

template<>
void cell::set_value(long double d)
{
    d_->is_date_ = false;
    d_->value_numeric_ = static_cast<long double>(d);
    d_->type_ = type::numeric;
}

template<>
void cell::set_value(std::string s)
{
    d_->set_string(s, get_parent().get_parent().get_guess_types());
}

template<>
void cell::set_value(char const *c)
{
    set_value(std::string(c));
}

template<>
void cell::set_value(cell c)
{
    d_->type_ = c.d_->type_;
    d_->value_numeric_ = c.d_->value_numeric_;
    d_->value_string_ = c.d_->value_string_;
    d_->is_date_ = c.d_->is_date_;
    d_->hyperlink_ = c.d_->hyperlink_;
    d_->has_hyperlink_ = c.d_->has_hyperlink_;
    d_->formula_ = c.d_->formula_;
    set_style(c.get_style());
    set_comment(c.get_comment());
}
    
template<>
void cell::set_value(date d)
{
    auto base_date = get_parent().get_parent().get_properties().excel_base_date;
    d_->set_date(d.to_number(base_date), xlnt::number_format::format::date_yyyymmdd2);
}

template<>
void cell::set_value(datetime d)
{
    auto base_date = get_parent().get_parent().get_properties().excel_base_date;
    d_->set_date(d.to_number(base_date), xlnt::number_format::format::date_datetime);
}

template<>
void cell::set_value(time t)
{
    d_->set_date(t.to_number(), xlnt::number_format::format::date_time6);
}

template<>
void cell::set_value(timedelta t)
{
    d_->set_date(t.to_number(), xlnt::number_format::format::date_timedelta);
    d_->is_date_ = false; // a timedelta isn't actually a date, still uses mostly the same code
}
    
bool cell::has_style() const
{
    return d_->style_ != nullptr;
}

row_t cell::get_row() const
{
    return d_->row_;
}

std::string cell::get_column() const
{
    return cell_reference::column_string_from_index(d_->column_);
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
    return d_->is_date_ || (d_->style_ != nullptr && get_style().get_number_format().get_format_code() == number_format::format::date_xlsx14);
}

cell_reference cell::get_reference() const
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

style &cell::get_style()
{
    return d_->get_style(true);
}

const style &cell::get_style() const
{
    return d_->get_style();
}
    
void cell::set_style(const xlnt::style &s)
{
    get_style() = s;
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

std::string cell::to_string() const
{
    return "<Cell " + worksheet(d_->parent_).get_title() + "." + get_reference().to_string() + ">";
}

relationship cell::get_hyperlink() const
{
    if(!d_->has_hyperlink_)
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
    if(hyperlink.length() == 0 || std::find(hyperlink.begin(), hyperlink.end(), ':') == hyperlink.end())
    {
        throw data_type_exception();
    }

    d_->has_hyperlink_ = true;
    d_->hyperlink_ = worksheet(d_->parent_).create_relationship(relationship::type::hyperlink, hyperlink);

    if(get_data_type() == type::null)
    {
        set_value(hyperlink);
    }
}

void cell::set_formula(const std::string &formula)
{
    if(formula.length() == 0)
    {
        throw data_type_exception();
    }

    d_->formula_ = formula;
}

bool cell::has_formula() const
{
    return !d_->formula_.empty();
}

std::string cell::get_formula() const
{
    if(d_->formula_.empty())
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
    if(c.d_ != d_->comment_.get())
    {
        throw xlnt::attribute_error();
    }
    
    if(!has_comment())
    {
        get_parent().increment_comments();
    }
    
    *get_comment().d_ = *c.d_;
}

void cell::clear_comment()
{
    if(has_comment())
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
    if(error.length() == 0 || error[0] != '#')
    {
        throw data_type_exception();
    }

    d_->value_string_ = error;
    d_->type_ = type::error;
}

cell cell::offset(column_t column, row_t row)
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
    if(d_->comment_ == nullptr)
    {
        d_->comment_.reset(new detail::comment_impl());
        get_parent().increment_comments();
    }
    
    return comment(d_->comment_.get());
}

std::pair<int, int> cell::get_anchor() const
{
    static const double DefaultColumnWidth = 51.85;
    static const double DefaultRowHeight = 15.0;

    auto points_to_pixels = [](double value, double dpi) { return (int)std::ceil(value * dpi / 72); };

    auto left_columns = d_->column_ - 1;
    auto &column_dimensions = get_parent().get_column_dimensions();
    int left_anchor = 0;
    auto default_width = points_to_pixels(DefaultColumnWidth, 96.0);

    for(column_t column_index = 1; column_index <= left_columns; column_index++)
    {
        if(column_dimensions.find(column_index) != column_dimensions.end())
        {
            auto cdw = column_dimensions.at(column_index);

            if(cdw > 0)
            {
                left_anchor += points_to_pixels(cdw, 96.0);
                continue;
            }
        }

        left_anchor += default_width;
    }

    auto top_rows = d_->row_ - 1;
    auto &row_dimensions = get_parent().get_row_dimensions();
    int top_anchor = 0;
    auto default_height = points_to_pixels(DefaultRowHeight, 96.0);

    for(int row_index = 1; row_index <= (int)top_rows; row_index++)
    {
        if(row_dimensions.find(static_cast<row_t>(row_index)) != row_dimensions.end())
        {
            auto rdh = row_dimensions.at(static_cast<row_t>(row_index));

            if(rdh > 0)
            {
                top_anchor += points_to_pixels(rdh, 96.0);
                continue;
            }
        }

        top_anchor += default_height;
    }

    return {left_anchor, top_anchor};
}

cell::type cell::get_data_type() const
{
    return d_->type_;
}

void cell::set_data_type(type t)
{
    d_->type_ = t;
}

std::size_t cell::get_xf_index() const
{
    return d_->xf_index_;
}

std::string cell::get_number_format()
{
    return get_style().get_number_format().get_format_code_string();
}
    
std::string cell::get_number_format() const
{
    return get_style().get_number_format().get_format_code_string();
}

font &cell::get_font()
{
    return get_style().get_font();
}
    
fill &cell::get_fill()
{
    return get_style().get_fill();
}

const fill &cell::get_fill() const
{
    return get_style().get_fill();
}

border &cell::get_border()
{
    return get_style().get_border();
}

alignment &cell::get_alignment()
{
    return get_style().get_alignment();
}

protection &cell::get_protection()
{
    return get_style().get_protection();
}

bool cell::pivot_button()
{
    return get_style().pivot_button();
}

bool cell::quote_prefix()
{
    return get_style().quote_prefix();
}
    
void cell::clear_value()
{
    d_->value_numeric_ = 0;
    d_->value_string_.clear();
    d_->formula_.clear();
    d_->type_ = cell::type::null;
}

template<>
bool cell::get_value() const
{
    return d_->value_numeric_ != 0;
}

template<>
std::int8_t cell::get_value() const
{
    return static_cast<std::int8_t>(d_->value_numeric_);
}

template<>
std::int16_t cell::get_value() const
{
    return static_cast<std::int16_t>(d_->value_numeric_);
}

template<>
std::int32_t cell::get_value() const
{
    return static_cast<std::int32_t>(d_->value_numeric_);
}

template<>
std::int64_t cell::get_value() const
{
    return static_cast<std::int64_t>(d_->value_numeric_);
}

template<>
std::uint8_t cell::get_value() const
{
    return static_cast<std::uint8_t>(d_->value_numeric_);
}

template<>
std::uint16_t cell::get_value() const
{
    return static_cast<std::uint16_t>(d_->value_numeric_);
}

template<>
std::uint32_t cell::get_value() const
{
    return static_cast<std::uint32_t>(d_->value_numeric_);
}

template<>
std::uint64_t cell::get_value() const
{
    return static_cast<std::uint64_t>(d_->value_numeric_);
}
    
template<>
float cell::get_value() const
{
    return static_cast<float>(d_->value_numeric_);
}

template<>
double cell::get_value() const
{
    return static_cast<double>(d_->value_numeric_);
}

template<>
long double cell::get_value() const
{
    return d_->value_numeric_;
}
    
template<>
time cell::get_value() const
{
    return time::from_number(d_->value_numeric_);
}

template<>
datetime cell::get_value() const
{
    return datetime::from_number(d_->value_numeric_, xlnt::calendar::windows_1900);
}

template<>
date cell::get_value() const
{
    return date::from_number(static_cast<int>(d_->value_numeric_), xlnt::calendar::windows_1900);
}

template<>
timedelta cell::get_value() const
{
    return timedelta(0, 0);
    //return timedelta::from_number(d_->value_numeric_);
}

void cell::set_number_format(const std::string &format_string)
{
    get_style().get_number_format().set_format_code_string(format_string);
}

template<>
std::string cell::get_value() const
{
    return d_->value_string_;
}

bool cell::has_value() const
{
    return d_->type_ != cell::type::null;
}

std::ostream &cell::print(std::ostream &stream, bool convert) const
{
    if(!convert)
    {
        return stream << get_value<std::string>();
    }
    else
    {
        switch(get_data_type())
        {
        case type::null:
            return stream << "";
        case type::string:
            return stream << get_value<std::string>();
        case type::numeric:
            if(is_date())
            {
                return stream << get_value<datetime>().to_string(get_parent().get_parent().get_properties().excel_base_date);
            }
            else
            {
                return stream << get_value<long double>();
            }
        case type::error:
            return stream << get_value<std::string>();
        case type::formula:
            return stream << d_->formula_;
        default:
            return stream;
        }
    }
}

} // namespace xlnt
