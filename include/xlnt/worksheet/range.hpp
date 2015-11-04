// Copyright (c) 2015 Thomas Fussell
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
#include <memory>
#include <string>
#include <vector>

#include <xlnt/worksheet/cell_vector.hpp>
#include <xlnt/worksheet/major_order.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>

#include "xlnt_config.hpp"

namespace xlnt {

class XLNT_CLASS range
{
  public:
    template <bool is_const = true>
    class common_iterator : public std::iterator<std::bidirectional_iterator_tag, cell>
    {
      public:
        common_iterator(worksheet ws, const range_reference &start_cell, major_order order = major_order::row)
            : ws_(ws), current_cell_(start_cell.get_top_left()), range_(start_cell), order_(order)
        {
        }

        common_iterator(const common_iterator<false> &other)
        {
            *this = other;
        }

        cell_vector operator*();

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
                current_cell_.set_row(current_cell_.get_row() - 1);
            }
            else
            {
                current_cell_.set_column_index(current_cell_.get_column_index() - 1);
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
                current_cell_.set_row(current_cell_.get_row() + 1);
            }
            else
            {
                current_cell_.set_column_index(current_cell_.get_column_index() + 1);
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

    range(worksheet ws, const range_reference &reference, major_order order = major_order::row);

    ~range();

    cell_vector operator[](std::size_t vector_index);

    const cell_vector operator[](std::size_t vector_index) const;

    bool operator==(const range &comparand) const;

    bool operator!=(const range &comparand) const
    {
        return !(*this == comparand);
    }

    cell_vector get_vector(std::size_t vector_index);

    const cell_vector get_vector(std::size_t vector_index) const;

    cell get_cell(const cell_reference &ref);

    const cell get_cell(const cell_reference &ref) const;

    range_reference get_reference() const;

    std::size_t length() const;

    bool contains(const cell_reference &ref);

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
