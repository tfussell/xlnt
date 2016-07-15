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

#include <xlnt/styles/fill.hpp>

namespace xlnt {

fill::type fill::get_type() const
{
    return type_;
}

pattern_fill::type pattern_fill::get_type() const
{
    return type_;
}

void pattern_fill::set_type(pattern_fill::type t)
{
    type_ = t;
}

gradient_fill::type gradient_fill::get_type() const
{
    return type_;
}

void gradient_fill::set_type(gradient_fill::type t)
{
    type_ = t;
}

void pattern_fill::set_foreground_color(const color &c)
{
    foreground_color_ = c;
}

std::experimental::optional<color> &pattern_fill::get_foreground_color()
{
    return foreground_color_;
}

const std::experimental::optional<color> &pattern_fill::get_foreground_color() const
{
    return foreground_color_;
}

void pattern_fill::set_background_color(const color &c)
{
    background_color_ = c;
}

std::experimental::optional<color> &pattern_fill::get_background_color()
{
    return background_color_;
}

const std::experimental::optional<color> &pattern_fill::get_background_color() const
{
    return background_color_;
}

void gradient_fill::set_degree(double degree)
{
    degree_ = degree;
}

double gradient_fill::get_degree() const
{
    return degree_;
}

std::string pattern_fill::to_hash_string() const
{
    std::string hash_string = "pattern_fill";

    hash_string.append(background_color_ ? "1" : "0");

    if (background_color_)
    {
        hash_string.append(std::to_string(background_color_->hash()));
    }

    hash_string.append(background_color_ ? "1" : "0");

    if (background_color_)
    {
        hash_string.append(std::to_string(background_color_->hash()));
    }
    
    return hash_string;
}

std::string gradient_fill::to_hash_string() const
{
    std::string hash_string = "gradient_fill";

    hash_string.append(std::to_string(static_cast<std::size_t>(type_)));
    hash_string.append(std::to_string(stops_.size()));
    hash_string.append(std::to_string(degree_));
    hash_string.append(std::to_string(left_));
    hash_string.append(std::to_string(right_));
    hash_string.append(std::to_string(top_));
    hash_string.append(std::to_string(bottom_));
    
    return hash_string;
}

std::string fill::to_hash_string() const
{
    std::string hash_string = "fill";
    
    hash_string.append(std::to_string(static_cast<std::size_t>(type_)));

    switch (type_)
    {
    case type::pattern:
        hash_string.append(std::to_string(pattern_.hash()));
        break;

    case type::gradient:
        hash_string.append(std::to_string(gradient_.hash()));
        break;

    default:
        break;
    } // switch (type_)

    return hash_string;
}

double gradient_fill::get_gradient_left() const
{
    return left_;
}

void gradient_fill::set_gradient_left(double value)
{
    left_ = value;
}

double gradient_fill::get_gradient_right() const
{
    return right_;
}

void gradient_fill::set_gradient_right(double value)
{
    right_ = value;
}

double gradient_fill::get_gradient_top() const
{
    return top_;
}

void gradient_fill::set_gradient_top(double value)
{
    top_ = value;
}

double gradient_fill::get_gradient_bottom() const
{
    return bottom_;
}

void gradient_fill::set_gradient_bottom(double value)
{
    bottom_ = value;
}

fill fill::pattern(pattern_fill::type pattern_type)
{
    fill f;
    f.type_ = type::pattern;
    f.pattern_ = pattern_fill(pattern_type);
    return f;
}

fill fill::gradient(gradient_fill::type gradient_type)
{
    fill f;
    f.type_ = type::gradient;
    f.gradient_ = gradient_fill(gradient_type);
    return f;
}

pattern_fill::pattern_fill(type pattern_type) : type_(pattern_type)
{

}

pattern_fill::pattern_fill() : pattern_fill(type::none)
{

}

gradient_fill::gradient_fill(type gradient_type) : type_(gradient_type)
{

}

gradient_fill::gradient_fill() : gradient_fill(type::linear)
{

}

gradient_fill &fill::get_gradient_fill()
{
    if (type_ != fill::type::gradient)
    {
        throw std::runtime_error("not gradient");
    }

    return gradient_;
}

const gradient_fill &fill::get_gradient_fill() const
{
    if (type_ != fill::type::gradient)
    {
        throw std::runtime_error("not gradient");
    }

    return gradient_;
}

pattern_fill &fill::get_pattern_fill()
{
    if (type_ != fill::type::pattern)
    {
        throw std::runtime_error("not pattern");
    }

    return pattern_;
}

const pattern_fill &fill::get_pattern_fill() const
{
    if (type_ != fill::type::pattern)
    {
        throw std::runtime_error("not pattern");
    }

    return pattern_;
}

void gradient_fill::add_stop(double position, color stop_color)
{
    stops_[position] = stop_color;
}

void gradient_fill::delete_stop(double position)
{
    stops_.erase(stops_.find(position));
}

void gradient_fill::clear_stops()
{
    stops_.clear();
}

const std::unordered_map<double, color> &gradient_fill::get_stops() const
{
    return stops_;
}

} // namespace xlnt
