// Copyright (c) 2014-2015 Thomas Fussell
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file
#pragma once

#include <iterator>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/worksheet/major_order.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

class cell;
class range_reference;

class XLNT_CLASS cell_vector
{
  public:
    template <bool is_const = true>
    class XLNT_CLASS common_iterator : public std::iterator<std::bidirectional_iterator_tag, cell>
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

    std::size_t length() const;

    iterator begin();
    iterator end();

    const_iterator begin() const;
    const_iterator cbegin() const;
    
    const_iterator end() const;
    const_iterator cend() const;

  private:
    worksheet ws_;
    range_reference ref_;
    major_order order_;
};

} // namespace xlnt
