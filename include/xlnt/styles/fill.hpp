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

#include <xlnt/xlnt_config.hpp>
#include <xlnt/styles/color.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

/// <summary>
/// Describes the fill style of a particular cell.
/// </summary>
class XLNT_CLASS fill : public hashable
{
  public:
    enum class type
    {
        none,
        solid,
        gradient,
        pattern
    };

    enum class gradient_type
    {
        linear,
        path
    };

    enum class pattern_type
    {
        none,
        solid,
        darkdown,
        darkgray,
        darkgrid,
        darkhorizontal,
        darktrellis,
        darkup,
        darkvertical,
        gray0625,
        gray125,
        lightdown,
        lightgray,
        lightgrid,
        lighthorizontal,
        lighttrellis,
        lightup,
        lightvertical,
        mediumgray,
    };

    type get_type() const;
    
    void set_type(type t);

    std::string get_pattern_type_string() const;
    
    std::string get_gradient_type_string() const;

    pattern_type get_pattern_type() const;

    void set_pattern_type(pattern_type t);

    void set_gradient_type(gradient_type t);

    std::experimental::optional<color> &get_foreground_color();
    
    const std::experimental::optional<color> &get_foreground_color() const;
    
    std::experimental::optional<color> &get_background_color();

    const std::experimental::optional<color> &get_background_color() const;
    
    std::experimental::optional<color> &get_start_color();
    
    const std::experimental::optional<color> &get_start_color() const;
    
    std::experimental::optional<color> &get_end_color();

    const std::experimental::optional<color> &get_end_color() const;

    void set_rotation(double rotation);

    double get_rotation() const;
    
    double get_gradient_left() const;

    double get_gradient_right() const;
    
    double get_gradient_top() const;
    
    double get_gradient_bottom() const;
    
protected:
    std::string to_hash_string() const override;

private:
    type type_ = type::none;
    pattern_type pattern_type_;
    gradient_type gradient_type_;
    double rotation_ = 0;
    std::experimental::optional<color> foreground_color_ = color::black();
    std::experimental::optional<color> background_color_ = color::white();
    std::experimental::optional<color> start_color_ = color::white();
    std::experimental::optional<color> end_color_ = color::black();
    double gradient_path_left_ = 0;
    double gradient_path_right_ = 0;
    double gradient_path_top_ = 0;
    double gradient_path_bottom_ = 0;
};

} // namespace xlnt
