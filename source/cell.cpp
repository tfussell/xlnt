#include <algorithm>
#include <locale>
#include <sstream>

#include "cell.h"
#include "cell_reference.h"
#include "cell_struct.h"
#include "relationship.h"
#include "worksheet.h"

namespace xlnt {

struct worksheet_struct;

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

cell::cell() : root_(nullptr)
{
}

cell::cell(worksheet &worksheet, const std::string &column, row_t row) : root_(nullptr)
{
    cell self = worksheet.get_cell(column + std::to_string(row));
    root_ = self.root_;
}


cell::cell(worksheet &worksheet, const std::string &column, row_t row, const std::string &initial_value) : root_(nullptr)
{
    cell self = worksheet.get_cell(column + std::to_string(row));
    root_ = self.root_;
    *this = initial_value;
}

cell::cell(cell_struct *root) : root_(root)
{
}

std::string cell::get_value() const
{
    switch(root_->type)
    {
        case type::string:
            return root_->string_value;
        case type::numeric:
            return std::to_string(root_->numeric_value);
        case type::formula:
            return root_->formula_value;
        case type::error:
            return root_->error_value;
        case type::null:
            return "";
        case type::date:
            return "00:00:00";
        case type::boolean:
            return root_->bool_value ? "TRUE" : "FALSE";
        default:
            throw std::runtime_error("bad enum");
    }
}

row_t cell::get_row() const
{
    return root_->row + 1;
}

std::string cell::get_column() const
{
    return cell_reference::column_string_from_index(root_->column + 1);
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
    root_->type = data_type;
    switch(data_type)
    {
        case type::formula: root_->formula_value = value; return;
        case type::date: root_->date_value.tm_hour = std::stoi(value); return;
        case type::error: root_->error_value = value; return;
        case type::boolean: root_->bool_value = value == "true"; return;
        case type::null: return;
        case type::numeric: root_->numeric_value = std::stod(value); return;
        case type::string: root_->string_value = value; return;
        default: throw std::runtime_error("bad enum");
    }
}

void cell::set_hyperlink(const std::string &url)
{
    root_->type = type::hyperlink;
    root_->hyperlink_rel = worksheet(root_->parent_worksheet).create_relationship("hyperlink", url);
}

void cell::set_merged(bool merged)
{
    root_->merged = merged;
}

bool cell::is_merged() const
{
    return root_->merged;
}

bool cell::bind_value()
{
    root_->type = type::null;
    return true;
}

bool cell::bind_value(int value)
{
    root_->type = type::numeric;
    root_->numeric_value = value;
    return true;
}

bool cell::bind_value(double value)
{
    root_->type = type::numeric;
    root_->numeric_value = value;
    return true;
}

bool cell::bind_value(const std::string &value)
{
    //Given a value, infer type and display options.
    root_->type = data_type_for_value(value);
    return true;
}

bool cell::bind_value(const char *value)
{
    return bind_value(std::string(value));
}

bool cell::bind_value(bool value)
{
    root_->type = type::boolean;
    root_->bool_value = value;
    return true;
}

bool cell::bind_value(const tm &value)
{
    root_->type = type::date;
    root_->date_value = value;
    return true;
}
    
bool cell::is_date() const
{
    return root_->type == type::date;
}

cell_reference cell::get_reference() const
{
    return {root_->column, root_->row};
}

std::string cell::get_hyperlink_rel_id() const
{
    return root_->hyperlink_rel.get_id();
}

bool cell::operator==(std::nullptr_t) const
{
    return root_ == nullptr;
}

bool cell::operator==(int comparand) const
{
    return root_->type == type::numeric && root_->numeric_value == comparand;
}

bool cell::operator==(double comparand) const
{
    return root_->type == type::numeric && root_->numeric_value == comparand;
}

bool cell::operator==(const std::string &comparand) const
{
    if(root_->type == type::hyperlink)
    {
        return root_->hyperlink_rel.get_target_uri() == comparand;
    }
    if(root_->type == type::string)
    {
        return root_->string_value == comparand;
    }
    return false;
}

bool cell::operator==(const char *comparand) const
{
    return *this == std::string(comparand);
}

bool cell::operator==(const tm &comparand) const
{
    return root_->type == cell::type::date && root_->date_value.tm_hour == comparand.tm_hour;
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
    if(root_->style_ == nullptr)
    {
        root_->style_ = new style();
    }
    return *root_->style_;
}

const style &cell::get_style() const
{
    if(root_->style_ == nullptr)
    {
        root_->style_ = new style();
    }
    return *root_->style_;
}

xlnt::cell::type cell::get_data_type() const
{
    return root_->type;
}

xlnt::cell cell::get_offset(int row_offset, int column_offset)
{
    worksheet parent_wrapper(root_->parent_worksheet);
    cell_reference ref(root_->column + column_offset, root_->row + row_offset);
    return parent_wrapper[ref];
}

cell &cell::operator=(const cell &rhs)
{
    root_ = rhs.root_;
    return *this;
}

cell &cell::operator=(int value)
{
    root_->type = type::numeric;
    root_->numeric_value = value;
    return *this;
}

cell &cell::operator=(double value)
{
    root_->type = type::numeric;
    root_->numeric_value = value;
    return *this;
}

cell &cell::operator=(bool value)
{
    root_->type = type::boolean;
    root_->bool_value = value;
    return *this;
}


cell &cell::operator=(const std::string &value)
{
    root_->type = data_type_for_value(value);
    
    switch(root_->type)
    {
        case type::date:
        {
            root_->date_value = std::tm();
            auto split = split_string(value, ':');
            root_->date_value.tm_hour = std::stoi(split[0]);
            root_->date_value.tm_min = std::stoi(split[1]);
            if(split.size() > 2)
            {
                root_->date_value.tm_sec = std::stoi(split[2]);
            }
            break;
        }
        case type::formula:
            root_->formula_value = value;
            break;
        case type::numeric:
            root_->numeric_value = std::stod(value);
            break;
        case type::boolean:
            root_->bool_value = value == "TRUE" || value == "true";
            break;
        case type::error:
            root_->error_value = value;
            break;
        case type::string:
            root_->string_value = value;
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
    root_->type = type::date;
    root_->date_value = value;
    return *this;
}

std::string cell::to_string() const
{
    return root_->to_string();
}
    
std::string cell_struct::to_string() const
{
    return "<Cell " + worksheet(parent_worksheet).get_title() + "." + cell_reference(column, row).to_string() + ">";
}

cell cell::allocate(xlnt::worksheet owner, column_t column_index, row_t row_index)
{
    return new cell_struct(owner.root_, column_index, row_index);
}

void cell::deallocate(xlnt::cell cell)
{
    delete cell.root_;
}

} // namespace xlnt

