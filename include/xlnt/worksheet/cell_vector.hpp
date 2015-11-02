#pragma once

#include <iterator>

#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/major_order.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

class cell;
class range_reference;

class cell_vector
{
  public:
    template <bool is_const = true>
    class common_iterator : public std::iterator<std::bidirectional_iterator_tag, cell>
    {
      public:
        common_iterator(worksheet ws, const cell_reference &start_cell, major_order order = major_order::row)
            : ws_(ws), current_cell_(start_cell), range_(start_cell.to_range()), order_(order)
        {
        }

        common_iterator(const common_iterator<false> &other)
        {
            *this = other;
        }

        cell operator*();

        bool operator==(const common_iterator &other) const
        {
            return ws_ == other.ws_ && current_cell_ == other.current_cell_ && order_ == other.order_;
        }

        bool operator!=(const common_iterator &other) const
        {
            return !(*this == other);
        }

        common_iterator &operator--()
        {
            if (order_ == major_order::row)
            {
                current_cell_.set_column_index(current_cell_.get_column_index() - 1);
            }
            else
            {
                current_cell_.set_row(current_cell_.get_row() - 1);
            }

            return *this;
        }

        common_iterator operator--(int)
        {
            iterator old = *this;
            --*this;
            return old;
        }

        common_iterator &operator++()
        {
            if (order_ == major_order::row)
            {
                current_cell_.set_column_index(current_cell_.get_column_index() + 1);
            }
            else
            {
                current_cell_.set_row(current_cell_.get_row() + 1);
            }

            return *this;
        }

        common_iterator operator++(int)
        {
            iterator old = *this;
            ++*this;
            return old;
        }

        friend class common_iterator<true>;

      private:
        worksheet ws_;
        cell_reference current_cell_;
        range_reference range_;
        major_order order_;
    };

    using iterator = common_iterator<false>;
    using const_iterator = common_iterator<true>;

    cell_vector(worksheet ws, const range_reference &ref, major_order order = major_order::row);

    std::size_t num_cells() const;

    cell front();

    const cell front() const;

    cell back();

    const cell back() const;

    cell operator[](std::size_t column_index);

    const cell operator[](std::size_t column_index) const;

    cell get_cell(std::size_t column_index);

    const cell get_cell(std::size_t column_index) const;

    std::size_t length() const
    {
        return num_cells();
    }

    iterator begin();
    iterator end();

    const_iterator begin() const
    {
        return cbegin();
    }
    const_iterator end() const
    {
        return cend();
    }

    const_iterator cbegin() const;
    const_iterator cend() const;

  private:
    worksheet ws_;
    range_reference ref_;
    major_order order_;
};

} // namespace xlnt
