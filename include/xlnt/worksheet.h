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

#include <set>
#include <string>
#include <vector>
#include <unordered_map>

namespace xlnt {

class cell;
class relationship;
class workbook;
struct worksheet_struct;

typedef std::vector<std::vector<cell>> range;

class worksheet
{
public:
    enum class Break
    {
        None = 0,
        Row = 1,
        Column = 2
    };

    enum class SheetState
    {
        Visible,
        Hidden,
        VeryHidden
    };

    enum class PaperSize
    {
        Letter = 1,
        LetterSmall = 2,
        Tabloid = 3,
        Ledger = 4,
        Legal = 5,
        Statement = 6,
        Executive = 7,
        A3 = 8,
        A4 = 9,
        A4Small = 10,
        A5 = 11
    };

    enum class Orientation
    {
        Portrait,
        Landscape
    };

    worksheet(workbook &parent_workbook, const std::string &title = "");
    void operator=(const worksheet &other);
    cell operator[](const std::string &address);
    std::string to_string() const;
    workbook &get_parent() const;
    void garbage_collect();
    std::set<cell> get_cell_collection();
    std::string get_title() const;
    void set_title(const std::string &title);
    cell get_freeze_panes() const;
    void set_freeze_panes(cell top_left_cell);
    void set_freeze_panes(const std::string &top_left_coordinate);
    void unfreeze_panes();
    xlnt::cell cell(const std::string &coordinate);
    xlnt::cell cell(int row, int column);
    int get_highest_row() const;
    int get_highest_column() const;
    std::string calculate_dimension() const;
    range range(const std::string &range_string, int row_offset, int column_offset);
    relationship create_relationship(const std::string &relationship_type);
    //void add_chart(chart chart);
    void merge_cells(const std::string &range_string);
    void merge_cells(int start_row, int start_column, int end_row, int end_column);
    void unmerge_cells(const std::string &range_string);
    void unmerge_cells(int start_row, int start_column, int end_row, int end_column);
    void append(const std::vector<xlnt::cell> &cells);
    void append(const std::unordered_map<std::string, xlnt::cell> &cells);
    void append(const std::unordered_map<int, xlnt::cell> &cells);
    xlnt::range rows() const;
    xlnt::range columns() const;
    bool operator==(const worksheet &other);

private:
    worksheet_struct *root_;
};

} // namespace xlnt
