/*
Copyright (c) 2012-2014 Thomas Fussell

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
*/
#pragma once

#include <ostream>
#include <string>

struct tm;

namespace xlnt {

class style;
class worksheet;
struct cell_struct;

struct coordinate
{
    std::string column;
    int row;
};

class cell
{
public:
    enum class type
    {
        null,
        numeric,
        string,
        date,
        formula,
        boolean,
        error
    };

    static coordinate coordinate_from_string(const std::string &address);
    static int column_index_from_string(const std::string &column_string);
    static std::string get_column_letter(int column_index);
    static std::string absolute_coordinate(const std::string &absolute_address);

    cell(const cell &copy);
    cell(worksheet &ws, const std::string &column, int row);
    cell(worksheet &ws, const std::string &column, int row, const std::string &value);
    ~cell();

    cell &operator=(int value);
    cell &operator=(double value);
    cell &operator=(const std::string &value);
    cell &operator=(const char *value);
    cell &operator=(const tm &value);

    friend bool operator==(const std::string &comparand, const cell &cell);
    friend bool operator==(const char *comparand, const cell &cell);
    friend bool operator==(const tm &comparand, const cell &cell);

    std::string to_string() const;
    bool is_date() const;
    style &get_style();
    type get_data_type();

private:
    cell_struct *root_;
};

inline std::ostream &operator<<(std::ostream &stream, const cell &cell)
{
    return stream << cell.to_string();
}

} // namespace xlnt
