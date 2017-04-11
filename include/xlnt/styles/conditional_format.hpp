// Copyright (c) 2014-2017 Thomas Fussell
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

#include <cstdint>
#include <string>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

class border;
class fill;
class font;

namespace detail {

struct conditional_format_impl;
struct stylesheet;
class xlsx_consumer;
class xlsx_producer;

} // namespace detail

class XLNT_API condition
{
public:
	static condition text_starts_with(const std::string &start);
	static condition text_ends_with(const std::string &end);
	static condition text_contains(const std::string &start);
	static condition text_does_not_contain(const std::string &start);

private:
	friend class detail::xlsx_producer;

	enum class type
	{
		contains_text
	} type_;

	enum class condition_operator
	{
		starts_with,
		ends_with,
		contains,
		does_not_contain
	} operator_;

	std::string text_comparand_;
};

/// <summary>
/// Describes a conditional format that will be applied to all cells in the 
/// associated range that satisfy the condition. This can only be constructed
/// using methods on worksheet or range.
/// </summary>
class XLNT_API conditional_format
{
public:
    /// <summary>
    /// Delete zero-argument constructor
    /// </summary>
	conditional_format() = delete;

    /// <summary>
    /// Default copy constructor. Constructs a format using the same PIMPL as other.
    /// </summary>
	conditional_format(const conditional_format &other) = default;

    // Formatting (xf) components

	/// <summary>
	///
	/// </summary>
	bool has_border() const;

    /// <summary>
    ///
    /// </summary>
    class border border() const;

    /// <summary>
    ///
    /// </summary>
	conditional_format border(const xlnt::border &new_border);

	/// <summary>
	///
	/// </summary>
	bool has_fill() const;

    /// <summary>
    ///
    /// </summary>
    class fill fill() const;

    /// <summary>
    ///
    /// </summary>
	conditional_format fill(const xlnt::fill &new_fill);

	/// <summary>
	///
	/// </summary>
	bool has_font() const;

    /// <summary>
    ///
    /// </summary>
    class font font() const;

    /// <summary>
    ///
    /// </summary>
	conditional_format font(const xlnt::font &new_font);

    /// <summary>
    /// Returns true if this format is equivalent to other.
    /// </summary>
    bool operator==(const conditional_format &other) const;

	/// <summary>
	/// Returns true if this format is not equivalent to other.
	/// </summary>
	bool operator!=(const conditional_format &other) const;

private:
    friend struct detail::stylesheet;
    friend class detail::xlsx_consumer;

    /// <summary>
    ///
    /// </summary>
	conditional_format(detail::conditional_format_impl *d);

    /// <summary>
    ///
    /// </summary>
    detail::conditional_format_impl *d_;
};

} // namespace xlnt
