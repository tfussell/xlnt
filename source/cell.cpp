#include <algorithm>
#include <locale>
#include <sstream>

#include "cell.h"
#include "cell_impl.h"
#include "cell_reference.h"
#include "relationship.h"
#include "worksheet.h"

namespace xlnt {

cell_impl::cell_impl() : hyperlink_rel("hyperlink"), column(0), row(0), type_(type::null)
{
    
}
    
cell_impl::cell_impl(int column_index, int row_index) : hyperlink_rel("hyperlink"), column(column_index), row(row_index), type_(type::null)
{
    
}
    
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

cell::cell(cell_impl *d) : d_(d)
{
    
}
    
cell::cell(worksheet worksheet, const cell_reference &reference, const std::string &initial_value) : d_(nullptr)
{
    cell self = worksheet.get_cell(reference);
    d_ = self.d_;

    if(initial_value != "")
    {
        *this = initial_value;
    }
}

std::string cell::get_value() const
{
    switch(d_->type_)
    {
        case cell_impl::type::string:
            return d_->string_value;
        case cell_impl::type::numeric:
            return std::to_string(d_->numeric_value);
        case cell_impl::type::formula:
            return d_->formula_value;
        case cell_impl::type::error:
            return d_->error_value;
        case cell_impl::type::null:
            return "";
        case cell_impl::type::date:
            return "00:00:00";
        case cell_impl::type::boolean:
            return d_->numeric_value != 0 ? "TRUE" : "FALSE";
        default:
            throw std::runtime_error("bad enum");
    }
}

row_t cell::get_row() const
{
    return d_->row + 1;
}

std::string cell::get_column() const
{
    return cell_reference::column_string_from_index(d_->column + 1);
}

std::vector<std::string> split_string(const std::string &string, char delim = ' ')
{
    std::stringstream ss(string);
    std::string part;
    std::vector<std::string> parts;
    while(std::getline(ss, part, delim))
    {
        parts.push_back(part);
    }
    return parts;
}

cell::type cell::data_type_for_value(const std::string &value)
{
    if(value.empty())
    {
        return type::null;
    }
    
    if(value[0] == '=')
    {
        return type::formula;
    }
    else if(value[0] == '0')
    {
        if(value.length() > 1)
        {
            if(value[1] == '.' || (value.length() > 2 && (value[1] == 'e' || value[1] == 'E')))
            {
                auto first_non_number = std::find_if(value.begin() + 2, value.end(),
                                                     [](char c) { return !std::isdigit(c, std::locale::classic()); });
                if(first_non_number == value.end())
                {
                    return type::numeric;
                }
            }
            auto split = split_string(value, ':');
            if(split.size() == 2 || split.size() == 3)
            {
                for(auto part : split)
                {
                    try
                    {
                        std::stoi(part);
                    }
                    catch(std::invalid_argument)
                    {
                        return type::string;
                    }
                }
                return type::date;
            }
            else
            {
                return type::string;
            }
        }
        return type::numeric;
    }
    else if(value[0] == '#')
    {
        return type::error;
    }
    else
    {
        char *p;
        strtod(value.c_str(), &p);
        if(*p != 0)
        {
            static const std::vector<std::string> possible_booleans = {"TRUE", "true", "FALSE", "false"};
            if(std::find(possible_booleans.begin(), possible_booleans.end(), value) != possible_booleans.end())
            {
                return type::boolean;
            }
            return type::string;
        }
        else
        {
            return type::numeric;
        }
    }
}

void cell::set_explicit_value(const std::string &value, type data_type)
{
    d_->type_ = (cell_impl::type)data_type;
    switch(data_type)
    {
        case type::formula: d_->formula_value = value; return;
        case type::date: d_->date_value.tm_hour = std::stoi(value); return;
        case type::error: d_->error_value = value; return;
        case type::boolean: d_->numeric_value = value == "true"; return;
        case type::null: return;
        case type::numeric: d_->numeric_value = std::stod(value); return;
        case type::string: d_->string_value = value; return;
        default: throw std::runtime_error("bad enum");
    }
}

void cell::set_hyperlink(const std::string &url)
{
    d_->type_ = cell_impl::type::hyperlink;
    d_->string_value = url;
}

void cell::set_merged(bool merged)
{
    d_->merged = merged;
}

bool cell::is_merged() const
{
    return d_->merged;
}

bool cell::bind_value()
{
    d_->type_ = cell_impl::type::null;
    return true;
}

bool cell::bind_value(int value)
{
    d_->type_ = cell_impl::type::numeric;
    d_->numeric_value = value;
    return true;
}

bool cell::bind_value(double value)
{
    d_->type_ = cell_impl::type::numeric;
    d_->numeric_value = value;
    return true;
}

bool cell::bind_value(const std::string &value)
{
    //Given a value, infer type and display options.
    d_->type_ = (cell_impl::type)data_type_for_value(value);
    return true;
}

bool cell::bind_value(const char *value)
{
    return bind_value(std::string(value));
}

bool cell::bind_value(bool value)
{
    d_->type_ = cell_impl::type::boolean;
    d_->numeric_value = value ? 1 : 0;
    return true;
}

bool cell::bind_value(const tm &value)
{
    d_->type_ = cell_impl::type::date;
    d_->date_value = value;
    return true;
}
    
bool cell::is_date() const
{
    return d_->type_ == cell_impl::type::date;
}

cell_reference cell::get_reference() const
{
    return {d_->column, d_->row};
}

std::string cell::get_hyperlink_rel_id() const
{
    return d_->hyperlink_rel.get_id();
}

bool cell::operator==(std::nullptr_t) const
{
    return d_ == nullptr;
}

bool cell::operator==(int comparand) const
{
    return d_->type_ == cell_impl::type::numeric && d_->numeric_value == comparand;
}

bool cell::operator==(double comparand) const
{
    return d_->type_ == cell_impl::type::numeric && d_->numeric_value == comparand;
}

bool cell::operator==(const std::string &comparand) const
{
    if(d_->type_ == cell_impl::type::hyperlink)
    {
        return d_->hyperlink_rel.get_target_uri() == comparand;
    }
    if(d_->type_ == cell_impl::type::string)
    {
        return d_->string_value == comparand;
    }
    return false;
}

bool cell::operator==(const char *comparand) const
{
    return *this == std::string(comparand);
}

bool cell::operator==(const tm &comparand) const
{
    return d_->type_ == cell_impl::type::date && d_->date_value.tm_hour == comparand.tm_hour;
}

bool operator==(int comparand, const xlnt::cell &cell)
{
    return cell == comparand;
}

bool operator==(const char *comparand, const xlnt::cell &cell)
{
    return cell == comparand;
}

bool operator==(const std::string &comparand, const xlnt::cell &cell)
{
    return cell == comparand;
}

bool operator==(const tm &comparand, const xlnt::cell &cell)
{
    return cell == comparand;
}

style &cell::get_style()
{
    if(d_->style_ == nullptr)
    {
        d_->style_ = new style();
    }
    return *d_->style_;
}

const style &cell::get_style() const
{
    if(d_->style_ == nullptr)
    {
        d_->style_ = new style();
    }
    return *d_->style_;
}

xlnt::cell::type cell::get_data_type() const
{
    return (type)d_->type_;
}

cell &cell::operator=(const cell &rhs)
{
    d_ = rhs.d_;
    return *this;
}

cell &cell::operator=(int value)
{
    d_->type_ = cell_impl::type::numeric;
    d_->numeric_value = value;
    return *this;
}

cell &cell::operator=(double value)
{
    d_->type_ = cell_impl::type::numeric;
    d_->numeric_value = value;
    return *this;
}

cell &cell::operator=(bool value)
{
    d_->type_ = cell_impl::type::boolean;
    d_->numeric_value = value ? 1 : 0;
    return *this;
}


cell &cell::operator=(const std::string &value)
{
    d_->type_ = (cell_impl::type)data_type_for_value(value);
    
    switch((type)d_->type_)
    {
        case type::date:
        {
            d_->date_value = std::tm();
            auto split = split_string(value, ':');
            d_->date_value.tm_hour = std::stoi(split[0]);
            d_->date_value.tm_min = std::stoi(split[1]);
            if(split.size() > 2)
            {
                d_->date_value.tm_sec = std::stoi(split[2]);
            }
            break;
        }
        case type::formula:
            d_->formula_value = value;
            break;
        case type::numeric:
            d_->numeric_value = std::stod(value);
            break;
        case type::boolean:
            d_->numeric_value = value == "TRUE" || value == "true";
            break;
        case type::error:
            d_->error_value = value;
            break;
        case type::string:
            d_->string_value = value;
            break;
        case type::null:
            break;
        default:
            throw std::runtime_error("bad enum");
    }
    
    return *this;
}

cell &cell::operator=(const char *value)
{
    return *this = std::string(value);
}

cell &cell::operator=(const tm &value)
{
    d_->type_ = cell_impl::type::date;
    d_->date_value = value;
    return *this;
}

std::string cell::to_string() const
{
    return "<Cell " + cell_reference(d_->column, d_->row).to_string() + ">";
}

} // namespace xlnt

