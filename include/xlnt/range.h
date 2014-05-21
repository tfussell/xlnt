#pragma once

#include <memory>
#include <string>
#include <vector>

#include "range_reference.h"
#include "worksheet.h"

namespace xlnt {
    
class row
{
public:
    row(worksheet ws, const range_reference &ref);

    std::size_t num_cells() const;

    cell front();

    const cell front() const;

    cell back();

    const cell back() const;

    cell operator[](std::size_t column_index);

    const cell operator[](std::size_t column_index) const;

    cell get_cell(std::size_t column_index);

    const cell get_cell(std::size_t column_index) const;

    class iterator
    {
    public:
        iterator(worksheet ws, const cell_reference &start_cell);
        ~iterator();

        bool operator==(const iterator &rhs);
        bool operator!=(const iterator &rhs) { return !(*this == rhs); }

        iterator operator++(int);
        iterator &operator++();

        cell operator*();

    private:
        worksheet ws_;
        cell_reference current_cell_;
        range_reference range_;
    };

    iterator begin();
    iterator end();

    class const_iterator
    {
    public:
        const_iterator(worksheet ws, const cell_reference &start_cell);
        ~const_iterator();

        bool operator==(const const_iterator &rhs);
        bool operator!=(const const_iterator &rhs) { return !(*this == rhs); }

        const_iterator operator++(int);
        const_iterator &operator++();

        const cell operator*();

    private:
        worksheet ws_;
        cell_reference current_cell_;
        range_reference range_;
    };

    const_iterator begin() const { return cbegin(); }
    const_iterator end() const { return cend(); }

    const_iterator cbegin() const;
    const_iterator cend() const;

private:
    worksheet ws_;
    range_reference ref_;
};

class range
{
public:
    range(worksheet ws, const range_reference &reference);

    ~range();

    row operator[](std::size_t row_index);

    const row operator[](std::size_t row_index) const;

    bool operator==(const range &comparand) const;

    bool operator!=(const range &comparand) const { return !(*this == comparand); }

    row get_row(std::size_t row_index);

    const row get_row(std::size_t row_index) const;

    cell get_cell(const cell_reference &ref);

    const cell get_cell(const cell_reference &ref) const;

    range_reference get_reference() const;

    std::size_t num_rows() const;

    class iterator
    {
    public:
        iterator(worksheet ws, const range_reference &start_cell);
        ~iterator();

        bool operator==(const iterator &rhs);
        bool operator!=(const iterator &rhs) { return !(*this == rhs); }

        iterator operator++(int);
        iterator &operator++();

        row operator*();

    private:
        worksheet ws_;
        cell_reference current_cell_;
        range_reference range_;
    };

    iterator begin();
    iterator end();

    class const_iterator
    {
    public:
        const_iterator(worksheet ws, const range_reference &start_cell);
        ~const_iterator();

        bool operator==(const const_iterator &rhs);
        bool operator!=(const const_iterator &rhs) { return !(*this == rhs); }

        const_iterator operator++(int);
        const_iterator &operator++();

        const row operator*();

    private:
        worksheet ws_;
        cell_reference current_cell_;
        range_reference range_;
    };

    const_iterator begin() const { return cbegin(); }
    const_iterator end() const { return cend(); }

    const_iterator cbegin() const;
    const_iterator cend() const;

private:
    worksheet ws_;
    range_reference ref_;
};

} // namespace xlnt
