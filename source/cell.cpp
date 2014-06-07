#include <algorithm>
#include <locale>
#include <sstream>

#include "cell/cell.hpp"
#include "cell/cell_reference.hpp"
#include "common/datetime.hpp"
#include "common/relationship.hpp"
#include "worksheet/worksheet.hpp"
#include "detail/cell_impl.hpp"
#include "common/exceptions.hpp"

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
    case type::string:
        return d_->string_value;
    case type::numeric:
        return std::to_string(d_->numeric_value);
    case type::formula:
        return d_->string_value;
    case type::error:
        return d_->string_value;
    case type::null:
        return "";
    case type::boolean:
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
                return type::numeric;
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
    d_->type_ = data_type;

    switch(data_type)
    {
        case type::null:
	  if(value != "")
	    {
	      throw data_type_exception();
	    }
	  return;
        case type::formula: 
	  if(value.length() == 0 || value[0] != '=')
	    {
	      throw data_type_exception();
	    }
	  d_->string_value = value; 
	  return;
        case type::error: d_->string_value = value; return;
        case type::boolean: d_->numeric_value = value == "true"; return;
        case type::numeric: d_->numeric_value = std::stod(value); return;
        case type::string: d_->string_value = value; return;
        default: throw std::runtime_error("bad enum");
    }
}

void cell::set_explicit_value(int value, type data_type)
{
    d_->type_ = data_type;

    switch(data_type)
    {
        case type::numeric: d_->numeric_value = value; return;
        case type::string: d_->string_value = std::to_string(value); return;
        default: throw data_type_exception();
    }
}

void cell::set_explicit_value(double value, type data_type)
{
    d_->type_ = data_type;

    switch(data_type)
    {
        case type::numeric: d_->numeric_value = value; return;
        case type::string: d_->string_value = std::to_string(value); return;
        default: throw data_type_exception();
    }
}

void cell::set_merged(bool merged)
{
    d_->merged = merged;
}

bool cell::is_merged() const
{
    return d_->merged;
}

bool cell::is_date() const
{
    return d_->is_date_;
}

cell_reference cell::get_reference() const
{
    return {d_->column, d_->row};
}

bool cell::operator==(std::nullptr_t) const
{
    return d_ == nullptr;
}

bool cell::operator==(int comparand) const
{
    return d_->type_ == type::numeric && d_->numeric_value == comparand;
}

bool cell::operator==(double comparand) const
{
    return d_->type_ == type::numeric && d_->numeric_value == comparand;
}

bool cell::operator==(const std::string &comparand) const
{
    if(d_->type_ == type::string)
    {
        return d_->string_value == comparand;
    }
    return false;
}

bool cell::operator==(const char *comparand) const
{
    return *this == std::string(comparand);
}

bool cell::operator==(const time &comparand) const
{
    return d_->type_ == type::numeric && time(d_->numeric_value).hour == comparand.hour;
}

bool cell::operator==(const date &comparand) const
{
    return d_->type_ == type::numeric && date((int)d_->numeric_value, 0, 0).year == comparand.year;
}

bool cell::operator==(const datetime &comparand) const
{
    return d_->type_ == type::numeric && datetime((int)d_->numeric_value, 0, 0).year == comparand.year;
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

bool operator==(const time &comparand, const xlnt::cell &cell)
{
    return cell == comparand;
}

bool operator==(const date &comparand, const xlnt::cell &cell)
{
    return cell == comparand;
}

bool operator==(const datetime &comparand, const xlnt::cell &cell)
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
    d_->is_date_ = false;
    d_->type_ = type::numeric;
    d_->numeric_value = value;
    return *this;
}

cell &cell::operator=(double value)
{
    d_->is_date_ = false;
    d_->type_ = type::numeric;
    d_->numeric_value = value;
    return *this;
}

cell &cell::operator=(long int value)
{
    d_->is_date_ = false;
    d_->type_ = type::numeric;
    d_->numeric_value = value;
    return *this;
}

cell &cell::operator=(long double value)
{
    d_->is_date_ = false;
    d_->type_ = type::numeric;
    d_->numeric_value = value;
    return *this;
}

cell &cell::operator=(bool value)
{
    d_->is_date_ = false;
    d_->type_ = type::boolean;
    d_->numeric_value = value ? 1 : 0;
    return *this;
}


cell &cell::operator=(const std::string &value)
{
    d_->is_date_ = false;
    d_->type_ = data_type_for_value(value);
    
    switch((type)d_->type_)
    {
        case type::formula:
            d_->string_value = value;
            break;
        case type::numeric:
	  if(value.find(':') != std::string::npos)
	    {
	      d_->is_date_ = true;
	      d_->numeric_value = time(value).to_number();
	    }
	  else
	    {
	      d_->numeric_value = std::stod(value);
	    }
            break;
        case type::boolean:
            d_->numeric_value = value == "TRUE" || value == "true";
            break;
        case type::error:
            d_->string_value = value;
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

cell &cell::operator=(const time &value)
{
    d_->type_ = type::numeric;
    d_->numeric_value = value.hour;
    return *this;
}

cell &cell::operator=(const date &value)
{
    d_->type_ = type::numeric;
    d_->numeric_value = value.year;
    d_->is_date_ = true;
    return *this;
}

cell &cell::operator=(const datetime &value)
{
    d_->type_ = type::numeric;
    d_->numeric_value = value.year;
    d_->is_date_ = true;
    return *this;
}

std::string cell::to_string() const
{
    return "<Cell " + worksheet(d_->parent_).get_title() + "." + get_reference().to_string() + ">";
}

  std::string cell::get_hyperlink() const
  {
    return "";
  }

  void cell::set_hyperlink(const std::string &hyperlink)
  {
    if(hyperlink.length() == 0 || std::find(hyperlink.begin(), hyperlink.end(), ':') == hyperlink.end())
      {
	throw data_type_exception();
      }
    d_->hyperlink_ = hyperlink;
    *this = hyperlink;
  }

  void cell::set_null()
  {
    d_->type_ = type::null;
  }

  void cell::set_formula(const std::string &formula)
  {
    if(formula.length() == 0 || formula[0] != '=')
      {
	throw data_type_exception();
      }
    d_->type_ = type::formula;
    d_->string_value = formula;
  }

  void cell::set_error(const std::string &error)
  {
    if(error.length() == 0 || error[0] != '#')
      {
	throw data_type_exception();
      }
    d_->type_ = type::error;
    d_->string_value = error;
  }

} // namespace xlnt
