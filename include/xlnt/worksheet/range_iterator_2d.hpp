#pragma once

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/worksheet/major_order.hpp>
#include <xlnt/worksheet/range_reference.hpp>

namespace xlnt {

class cell_vector;
class worksheet;

namespace detail { struct worksheet_impl; }

/// <summary>
/// An iterator used by worksheet and range for traversing
/// a 2D grid of cells by row/column then across that row/column.
/// </summary>
class XLNT_CLASS range_iterator_2d : public std::iterator<std::bidirectional_iterator_tag, cell_vector>
{
  public:
    range_iterator_2d(worksheet &ws, const range_reference &start_cell, major_order order = major_order::row);

    range_iterator_2d(const range_iterator_2d &other);

    cell_vector operator*();

    bool operator==(const range_iterator_2d &other) const;

    bool operator!=(const range_iterator_2d &other) const;

    range_iterator_2d &operator--();

    range_iterator_2d operator--(int);

    range_iterator_2d &operator++();

    range_iterator_2d operator++(int);

  private:
    detail::worksheet_impl *ws_;
    cell_reference current_cell_;
    range_reference range_;
    major_order order_;
};


/// <summary>
/// A const version of range_iterator_2d which does not allow modification
/// to the dereferenced cell_vector.
/// </summary>
class XLNT_CLASS const_range_iterator_2d : public std::iterator<std::bidirectional_iterator_tag, const cell_vector>
{
  public:
    const_range_iterator_2d(const worksheet &ws, const range_reference &start_cell, major_order order = major_order::row);

    const_range_iterator_2d(const const_range_iterator_2d &other);

    const cell_vector operator*();

    bool operator==(const const_range_iterator_2d &other) const;

    bool operator!=(const const_range_iterator_2d &other) const;

    const_range_iterator_2d &operator--();

    const_range_iterator_2d operator--(int);

    const_range_iterator_2d &operator++();

    const_range_iterator_2d operator++(int);

  private:
    detail::worksheet_impl *ws_;
    cell_reference current_cell_;
    range_reference range_;
    major_order order_;
};

} // namespace xlnt
