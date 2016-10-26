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

#include <cstdint>
#include <string>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/styles/base_format.hpp>

namespace xlnt {

namespace detail {
struct style_impl;
struct stylesheet;
}

/// <summary>
/// Describes a style which has a name and can be applied to multiple individual
/// formats. In Excel this is a "Cell Style".
/// </summary>
class XLNT_API style : public base_format
{
public:
    std::string name() const;
    style &name(const std::string &name);
    
    bool hidden() const;
	style &hidden(bool value);

	optional<bool> custom() const;
	style &custom(bool value);
    
    optional<std::size_t> builtin_id() const;
	style &builtin_id(std::size_t builtin_id);

	style &alignment_id(std::size_t id);
	style &border_id(std::size_t id);
	style &fill_id(std::size_t id);
	style &font_id(std::size_t id);
	style &number_format_id(std::size_t id);
	style &protection_id(std::size_t id);

	bool operator==(const style &other);

private:
	friend struct detail::stylesheet;
	style(detail::style_impl *d);
	detail::style_impl *d_;
};

} // namespace xlnt
