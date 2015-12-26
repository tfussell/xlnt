// Copyright (c) 2014-2016 Thomas Fussell
// Copyright (c) 2010-2015 openpyxl
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

#include <cstddef>
#include <functional>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/styles/diagonal_direction.hpp>
#include <xlnt/styles/side.hpp>
#include <xlnt/utils/hashable.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

/// <summary>
/// Describes the border style of a particular cell.
/// </summary>
class XLNT_CLASS border : public hashable
{
public:
    static border default_border();
    
    std::experimental::optional<side> &get_start();
    const std::experimental::optional<side> &get_start() const;
    std::experimental::optional<side> &get_end();
    const std::experimental::optional<side> &get_end() const;
    std::experimental::optional<side> &get_left();
    const std::experimental::optional<side> &get_left() const;
    std::experimental::optional<side> &get_right();
    const std::experimental::optional<side> &get_right() const;
    std::experimental::optional<side> &get_top();
    const std::experimental::optional<side> &get_top() const;
    std::experimental::optional<side> &get_bottom();
    const std::experimental::optional<side> &get_bottom() const;
    std::experimental::optional<side> &get_diagonal();
    const std::experimental::optional<side> &get_diagonal() const;
    std::experimental::optional<side> &get_vertical();
    const std::experimental::optional<side> &get_vertical() const;
    std::experimental::optional<side> &get_horizontal();
    const std::experimental::optional<side> &get_horizontal() const;

protected:
    std::string to_hash_string() const override;
    
private:
    std::experimental::optional<side> start_;
    std::experimental::optional<side> end_;
    std::experimental::optional<side> left_;
    std::experimental::optional<side> right_;
    std::experimental::optional<side> top_;
    std::experimental::optional<side> bottom_;
    std::experimental::optional<side> diagonal_;
    std::experimental::optional<side> vertical_;
    std::experimental::optional<side> horizontal_;

    bool outline_ = false;
    bool diagonal_up_ = false;
    bool diagonal_down_ = false;

    diagonal_direction diagonal_direction_ = diagonal_direction::none;
};

} // namespace xlnt
