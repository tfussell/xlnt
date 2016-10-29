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

// pattern_fill

pattern_fill::pattern_fill()
{

}

pattern_fill_type pattern_fill::type() const
{
	return type_;
}

pattern_fill &pattern_fill::type(pattern_fill_type type)
{
	type_ = type;
	return *this;
}

optional<color> pattern_fill::foreground() const
{
	return foreground_;
}

pattern_fill &pattern_fill::foreground(const color &new_foreground)
{
	foreground_ = new_foreground;

	if (!background_)
	{
		background_.set(indexed_color(64));
	}

	return *this;
}

optional<color> pattern_fill::background() const
{
	return background_;
}

pattern_fill &pattern_fill::background(const color &new_background)
{
	background_ = new_background;
	return *this;
}

std::string pattern_fill::to_hash_string() const
{
	std::string hash_string = "pattern_fill";

	hash_string.append(background_ ? "1" : "0");

	if (background_)
	{
		hash_string.append(std::to_string(background_->hash()));
	}

	hash_string.append(background_ ? "1" : "0");

	if (background_)
	{
		hash_string.append(std::to_string(background_->hash()));
	}

	return hash_string;
}

// gradient_fill

gradient_fill::gradient_fill() : type_(gradient_fill_type::linear)
{
}

gradient_fill_type gradient_fill::type() const
{
	return type_;
}

gradient_fill &gradient_fill::type(gradient_fill_type t)
{
	type_ = t;
	return *this;
}

gradient_fill &gradient_fill::degree(double degree)
{
	degree_ = degree;
	return *this;
}

double gradient_fill::degree() const
{
	return degree_;
}


double gradient_fill::left() const
{
	return left_;
}

gradient_fill &gradient_fill::left(double value)
{
	left_ = value;
	return *this;
}

double gradient_fill::right() const
{
	return right_;
}

gradient_fill &gradient_fill::right(double value)
{
	right_ = value;
	return *this;
}

double gradient_fill::top() const
{
	return top_;
}

gradient_fill &gradient_fill::top(double value)
{
	top_ = value;
	return *this;
}

double gradient_fill::bottom() const
{
	return bottom_;
}

gradient_fill &gradient_fill::bottom(double value)
{
	bottom_ = value;
	return *this;
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

gradient_fill &gradient_fill::add_stop(double position, color stop_color)
{
	stops_[position] = stop_color;
	return *this;
}

gradient_fill &gradient_fill::clear_stops()
{
	stops_.clear();
	return *this;
}

std::unordered_map<double, color> gradient_fill::stops() const
{
	return stops_;
}

// fill

fill fill::solid(const color &fill_color)
{
	return fill(xlnt::pattern_fill()
		.type(xlnt::pattern_fill_type::solid)
		.foreground(fill_color)
		.background(indexed_color(64)));
}

fill::fill() : type_(fill_type::pattern)
{
}

fill::fill(const xlnt::pattern_fill &pattern) 
	: type_(fill_type::pattern),
	  pattern_(pattern)
{
}

fill::fill(const xlnt::gradient_fill &gradient) 
	: type_(fill_type::gradient),
	  gradient_(gradient)
{
}

fill_type fill::type() const
{
    return type_;
}

gradient_fill fill::gradient_fill() const
{
    if (type_ != fill_type::gradient)
    {
		throw invalid_attribute();
    }

    return gradient_;
}

pattern_fill fill::pattern_fill() const
{
    if (type_ != fill_type::pattern)
    {
		throw invalid_attribute();
    }

    return pattern_;
}

std::string fill::to_hash_string() const
{
	std::string hash_string = "fill";

	hash_string.append(std::to_string(static_cast<std::size_t>(type_)));

	switch (type_)
	{
	case fill_type::pattern:
		hash_string.append(std::to_string(pattern_.hash()));
		break;

	case fill_type::gradient:
		hash_string.append(std::to_string(gradient_.hash()));
		break;

	default:
		break;
	} // switch (type_)

	return hash_string;
}

} // namespace xlnt
