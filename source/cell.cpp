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

namespace {

enum class condition_type
{
    less_than,
    less_or_equal,
    equal,
    greater_than,
    greater_or_equal,
    invalid
};
    
struct section
{
    bool has_value = false;
    std::string value;
    bool has_color = false;
    std::string color;
    bool has_condition = false;
    condition_type condition = condition_type::invalid;
    std::string condition_value;
    
    section &operator=(const section &other)
    {
        has_value = other.has_value;
        value = other.value;
        has_color = other.has_color;
        color = other.color;
        has_condition = other.has_condition;
        condition = other.condition;
        condition_value = other.condition_value;
        
        return *this;
    }
};
    
struct format_sections
{
    section first;
    section second;
    section third;
    section fourth;
};

// copied from named_range.cpp, keep in sync
/// <summary>
/// Return a vector containing string split at each delim.
/// </summary>
/// <remark>
/// This should maybe be in a utility header so it can be used elsewhere.
/// </remarks>
std::vector<std::string> split_string(const std::string &string, char delim)
{
    std::vector<std::string> split;
    std::string::size_type previous_index = 0;
    auto separator_index = string.find(delim);
    
    while(separator_index != std::string::npos)
    {
        auto part = string.substr(previous_index, separator_index - previous_index);
        split.push_back(part);
        
        previous_index = separator_index + 1;
        separator_index = string.find(delim, previous_index);
    }
    
    split.push_back(string.substr(previous_index));
    
    return split;
}

bool is_valid_color(const std::string &color)
{
    static const std::vector<std::string> colors = { "Black", "Green"
        "White", "Blue", "Magenta", "Yellow", "Cyan", "Red" };
    return std::find(colors.begin(), colors.end(), color) != colors.end();
}
    
bool parse_condition(const std::string &string, section &s)
{
    s.has_condition = false;
    s.condition = condition_type::invalid;
    s.condition_value.clear();
    
    if(string[0] == '<')
    {
        s.has_condition = true;
        
        if(string[1] == '=')
        {
            s.condition = condition_type::less_or_equal;
            s.condition_value = string.substr(2);
        }
        else
        {
            s.condition = condition_type::less_than;
            s.condition_value = string.substr(1);
        }
    }
    if(string[0] == '>')
    {
        s.has_condition = true;
        
        if(string[1] == '=')
        {
            s.condition = condition_type::greater_or_equal;
            s.condition_value = string.substr(2);
        }
        else
        {
            s.condition = condition_type::greater_than;
            s.condition_value = string.substr(1);
        }
    }
    else if(string[0] == '=')
    {
        s.has_condition = true;
        s.condition = condition_type::equal;
        s.condition_value = string.substr(1);
    }
    
    return s.has_condition;
}

section parse_section(const std::string &section_string)
{
    std::string intermediate = section_string;
    
    section s;
    std::string first_bracket_part, second_bracket_part;
    static const std::vector<std::string> bracket_times = { "h", "hh", "m", "mm", "s", "ss" };
    
    if(section_string[0] == '[')
    {
        auto close_bracket_pos = section_string.find(']');
        
        if(close_bracket_pos == std::string::npos)
        {
            throw std::runtime_error("missing close bracket");
        }
        
        first_bracket_part = intermediate.substr(1, close_bracket_pos - 1);
        
        if(std::find(bracket_times.begin(), bracket_times.end(), first_bracket_part) !=  bracket_times.end())
        {
            first_bracket_part.clear();
        }
        else
        {
            intermediate = intermediate.substr(close_bracket_pos);
        }
    }
    
    if(!first_bracket_part.empty() && intermediate[0] == '[')
    {
        auto close_bracket_pos = section_string.find(']');
        
        if(close_bracket_pos == std::string::npos)
        {
            throw std::runtime_error("missing close bracket");
        }
        
        second_bracket_part = intermediate.substr(1, close_bracket_pos - 1);
        
        if(std::find(bracket_times.begin(), bracket_times.end(), second_bracket_part) !=  bracket_times.end())
        {
            second_bracket_part.clear();
        }
        else
        {
            intermediate = intermediate.substr(close_bracket_pos);
        }
    }
    
    if(!first_bracket_part.empty())
    {
        if(is_valid_color(first_bracket_part))
        {
            s.color = first_bracket_part;
            s.has_color = true;
        }
        else if(!parse_condition(first_bracket_part, s))
        {
            throw std::runtime_error("invalid condition");
        }
    }
    
    if(!second_bracket_part.empty())
    {
        if(is_valid_color(second_bracket_part))
        {
            if(s.has_color)
            {
                throw std::runtime_error("two colors in one section");
            }
            
            s.color = second_bracket_part;
            s.has_color = true;
        }
        else if(s.has_condition)
        {
            throw std::runtime_error("two conditions in one section");
        }
        else if(!parse_condition(second_bracket_part, s))
        {
            throw std::runtime_error("invalid condition");
        }
    }
    
    s.value = intermediate;
    s.has_value = true;
    
    return s;
}
    
format_sections parse_format_sections(const std::string &combined)
{
    format_sections result = {};
    
    auto split = split_string(combined, ';');
    
    if(split.empty())
    {
        throw std::runtime_error("empty string");
    }
    
    result.first = parse_section(split[0]);
    
    if(!result.first.has_condition)
    {
        result.second = result.first;
        result.third = result.first;
    }
    
    if(split.size() > 1)
    {
        result.second = parse_section(split[1]);
    }
    
    if(split.size() > 2)
    {
        if(result.first.has_condition && !result.second.has_condition)
        {
            throw std::runtime_error("first two sections should have conditions");
        }
        
        result.third = parse_section(split[2]);
        
        if(result.third.has_condition)
        {
            throw std::runtime_error("third section shouldn't have a condition");
        }
    }
    
    if(split.size() > 3)
    {
        if(result.first.has_condition)
        {
            throw std::runtime_error("too many parts");
        }
        
        result.fourth = parse_section(split[3]);
    }
    
    if(split.size() > 4)
    {
        throw std::runtime_error("too many parts");
    }
    
    return result;
}

std::string format_section(long double number, const section &format)
{
    if(number == static_cast<long long int>(number))
    {
        return std::to_string(static_cast<long long int>(number));
    }
    
    return std::to_string(number);
}

std::string format_section(const std::string &text, const section &format)
{
    auto arobase_index = format.value.find('@');
    
    std::string first_part, middle_part, last_part;
    
    if(arobase_index != std::string::npos)
    {
        first_part = format.value.substr(0, arobase_index);
        middle_part = text;
        last_part = format.value.substr(arobase_index + 1);
    }
    else
    {
        first_part = format.value;
    }
    
    auto unquote = [](std::string &s)
    {
        if(!s.empty())
        {
            if(s.front() != '"' || s.back() != '"') return false;
            s = s.substr(0, s.size() - 2);
        }
        
        return true;
    };
    
    if(!unquote(first_part) || !unquote(last_part))
    {
        throw std::runtime_error(std::string("additional text must be enclosed in quotes: ") + format.value);
    }
    
    return first_part + middle_part + last_part;
}

std::string format_number(long double number, const std::string &format)
{
    auto sections = parse_format_sections(format);
    
    if(number > 0)
    {
        return format_section(number, sections.first);
    }
    else if(number < 0)
    {
        return format_section(number, sections.second);
    }
    
    // number == 0
    return format_section(number, sections.third);
}

std::string format_text(const std::string &text, const std::string &format)
{
    if(format == "General") return text;
    auto sections = parse_format_sections(format);
    return format_section(text, sections.fourth);
}

bool is_date_format(const std::string &format_string)
{
    auto not_in = format_string.find_first_not_of("/-:, mMyYdDhHsS");
    return not_in == std::string::npos;
}

const std::string PercentRegex("^\\-?(?P<number>[0-9]*\\.?[0-9]*\\s?)\%$");
const std::string TimeRegex("^(([0-1]{0,1}[0-9]{2}):([0-5][0-9]):?([0-5][0-9])?$)|"
                            "^(([0-5][0-9]):([0-5][0-9])?\\.(\\d{1,6}))");
const std::string NumberRegex("^-?([\\d]|[\\d]+\\.[\\d]*|\\.[\\d]+|[1-9][\\d]+\\.?[\\d]*)((E|e)[-+]?[\\d]+)?$");
    
}

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
    if(get_data_type() == type::numeric && has_style())
    {
        auto number_format = get_style().get_number_format().get_format_code_string();
        
        if(number_format != "General")
        {
            try
            {
                auto sections = parse_format_sections(number_format);
                return is_date_format(sections.first.value);
            }
            catch(std::exception)
            {
                return false;
            }
        }
    }
    
    return false;
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

std::string cell::to_repr() const
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

void cell::set_number_format(const std::string &format_string, int index)
{
    get_style().get_number_format().set_format_code_string(format_string, index);
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

std::string cell::to_string() const
{
    std::string number_format = "General";
    
    if(has_style())
    {
        number_format = get_style().get_number_format().get_format_code_string();
    }
    
    switch(get_data_type())
    {
        case cell::type::null:
            return "";
        case cell::type::numeric:
            return format_number(get_value<long double>(), number_format);
        case cell::type::string:
        case cell::type::formula:
        case cell::type::error:
            return format_text(get_value<std::string>(), number_format);
        case cell::type::boolean:
            return get_value<long double>() == 0 ? "FALSE" : "TRUE";
    }
}

} // namespace xlnt
