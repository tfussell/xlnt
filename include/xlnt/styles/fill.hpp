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

#include <unordered_map>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/styles/color.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

class pattern_fill : public hashable
{
public:
    enum class type
    {
        none,
        solid,
        mediumgray,
        darkgray,
        lightgray,
        darkhorizontal,
        darkvertical,
        darkdown,
        darkup,
        darkgrid,
        darktrellis,
        lighthorizontal,
        lightvertical,
        lightdown,
        lightup,
        lightgrid,
        lighttrellis,
        gray125,
        gray0625
    };

    pattern_fill();

    pattern_fill(type pattern_type);

    type get_type() const;

    void set_type(type pattern_type);

    void set_foreground_color(const color &c);

    std::experimental::optional<color> &get_foreground_color();

    const std::experimental::optional<color> &get_foreground_color() const;
    
    void set_background_color(const color &c);

    std::experimental::optional<color> &get_background_color();

    const std::experimental::optional<color> &get_background_color() const;

protected:
    std::string to_hash_string() const override;

private:
    type type_ = type::none;

    std::experimental::optional<color> foreground_color_;
    std::experimental::optional<color> background_color_;
};

class gradient_fill : public hashable
{
public:
    enum class type
    {
        linear,
        path
    };

    gradient_fill();

    gradient_fill(type gradient_type);

    type get_type() const;

    void set_type(type gradient_type);

    void set_degree(double degree);

    double get_degree() const;

    double get_gradient_left() const;

    void set_gradient_left(double value);

    double get_gradient_right() const;

    void set_gradient_right(double value);

    double get_gradient_top() const;

    void set_gradient_top(double value);

    double get_gradient_bottom() const;

    void set_gradient_bottom(double value);

    void add_stop(double position, color stop_color);

    void delete_stop(double position);

    void clear_stops();
    
    const std::unordered_map<double, color> &get_stops() const;

protected:
    std::string to_hash_string() const override;

private:
    type type_ = type::linear;

    std::unordered_map<double, color> stops_;

    double degree_ = 0;

    double left_ = 0;
    double right_ = 0;
    double top_ = 0;
    double bottom_ = 0;
};

/// <summary>
/// Describes the fill style of a particular cell.
/// </summary>
class XLNT_CLASS fill : public hashable
{
public:
    enum class type
    {
        pattern,
        gradient
    };

    static fill gradient(gradient_fill::type gradient_type);

    static fill pattern(pattern_fill::type pattern_type);

    type get_type() const;

    gradient_fill &get_gradient_fill();

    const gradient_fill &get_gradient_fill() const;

    pattern_fill &get_pattern_fill();

    const pattern_fill &get_pattern_fill() const;

protected:
    std::string to_hash_string() const override;

private:
    type type_ = type::pattern;
    gradient_fill gradient_;
    pattern_fill pattern_;
};

} // namespace xlnt
