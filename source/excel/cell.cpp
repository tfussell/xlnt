#include <ctime>

#include "cell.h"
#include "worksheet.h"

namespace xlnt {

struct cell_struct
{
    cell_struct() : type(cell::type::null)
    {

    }

    cell_struct(worksheet *ws, const std::string &column, int row) 
        : type(cell::type::null), parent_worksheet(ws), 
        column(xlnt::cell::column_index_from_string(column) - 1), row(row)
    {

    }

    std::string to_string() const
    {
        switch(type)
        {
        case cell::type::numeric: return std::to_string(numeric_value);
        case cell::type::boolean: return bool_value ? "true" : "false";
        case cell::type::string: return string_value;
        default: throw std::runtime_error("bad enum");
        }
    }

    cell::type type;

    union
    {
        long double numeric_value;
        bool bool_value;
    };

    tm date_value;
    std::string string_value;
    worksheet *parent_worksheet;
    int column;
    int row;
};

cell::cell(const cell &copy)
{
    root_ = copy.root_;
}

coordinate cell::coordinate_from_string(const std::string &address)
{
    if(address == "A1")
    {
        return {"A", 1};
    }

    return {"A", 1};
}

int cell::column_index_from_string(const std::string &column_string)
{
    return column_string[0] - 'A';
}

std::string cell::get_column_letter(int column_index)
{
    return std::string(1, (char)('A' + column_index - 1));
}

bool cell::is_date() const
{
    return root_->type == type::date;
}

bool operator==(const char *comparand, const cell &cell)
{
    return std::string(comparand) == cell;
}

bool operator==(const std::string &comparand, const cell &cell)
{
    return cell.root_->type == cell::type::string && cell.root_->string_value == comparand;
}

bool operator==(const tm &comparand, const cell &cell)
{
    return cell.root_->type == cell::type::date && cell.root_->date_value.tm_hour == comparand.tm_hour;
}

std::string cell::absolute_coordinate(const std::string &absolute_address)
{
    return absolute_address;
}

cell::cell(worksheet &ws, const std::string &column, int row) : root_(new cell_struct(&ws, column, row))
{

}

cell::cell(worksheet &ws, const std::string &column, int row, const std::string &value) : root_(new cell_struct(&ws, column, row))
{
    *this = value;
}

cell::type cell::get_data_type()
{
    return root_->type;
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

cell &cell::operator=(const std::string &value)
{
    root_->type = type::string;
    root_->string_value = value;
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

cell::~cell()
{
    delete root_;
}

std::string cell::to_string() const
{
    return root_->to_string();
}

} // namespace xlnt
