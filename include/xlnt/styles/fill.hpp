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

enum class XLNT_API pattern_fill_type
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

class XLNT_API pattern_fill : public hashable
{
public:
    pattern_fill();

	pattern_fill_type type() const;

	pattern_fill &type(pattern_fill_type new_type);

	optional<color> foreground() const;

	pattern_fill &foreground(const color &foreground);
    
    optional<color> background() const;

	pattern_fill &background(const color &background);

protected:
    std::string to_hash_string() const override;

private:
	pattern_fill_type type_ = pattern_fill_type::none;

    optional<color> foreground_;
    optional<color> background_;
};

enum class XLNT_API gradient_fill_type
{
	linear,
	path
};

class XLNT_API gradient_fill : public hashable
{
public:
    gradient_fill();

    gradient_fill_type type() const;

	// Type
    gradient_fill &type(gradient_fill_type new_type);

	// Degree

    gradient_fill &degree(double degree);

    double degree() const;

	// Left

    double left() const;

	gradient_fill &left(double value);

	// Right

    double right() const;

	gradient_fill &right(double value);

	// Top

    double top() const;

	gradient_fill &top(double value);

	// Bottom

    double bottom() const;

	gradient_fill &bottom(double value);

	// Stops

	gradient_fill &add_stop(double position, color stop_color);

	gradient_fill &clear_stops();
    
    std::unordered_map<double, color> stops() const;

protected:
    std::string to_hash_string() const override;

private:
    gradient_fill_type type_ = gradient_fill_type::linear;

    double degree_ = 0;

    double left_ = 0;
    double right_ = 0;
    double top_ = 0;
    double bottom_ = 0;

	std::unordered_map<double, color> stops_;
};

enum class XLNT_API fill_type
{
	pattern,
	gradient
};

/// <summary>
/// Describes the fill style of a particular cell.
/// </summary>
class XLNT_API fill : public hashable
{
public:
	/// <summary>
	/// Constructs a fill initialized as a none-type pattern fill with no 
	/// foreground or background colors.
	/// </summary>
	fill();

	/// <summary>
	/// Constructs a fill initialized as a pattern fill based on the given pattern.
	/// </summary>
	fill(const pattern_fill &pattern);

	/// <summary>
	/// Constructs a fill initialized as a gradient fill based on the given gradient.
	/// </summary>
	fill(const gradient_fill &gradient);

	/// <summary>
	/// Returns the fill_type of this fill depending on how it was constructed.
	/// </summary>
    fill_type type() const;

	/// <summary>
	/// Returns the gradient fill represented by this fill.
	/// Throws an invalid_attribute exception if this is not a gradient fill.
	/// </summary>
    class gradient_fill gradient_fill() const;

	/// <summary>
	/// Returns the pattern fill represented by this fill.
	/// Throws an invalid_attribute exception if this is not a pattern fill.
	/// </summary>
    class pattern_fill pattern_fill() const;

protected:
    std::string to_hash_string() const override;

private:
    fill_type type_ = fill_type::pattern;
    xlnt::gradient_fill gradient_;
    xlnt::pattern_fill pattern_;
};

} // namespace xlnt
