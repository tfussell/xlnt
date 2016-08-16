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

#include <xlnt/styles/font.hpp>

namespace xlnt {


font::font()
    : name_("Calibri"),
      size_(12),
      bold_(false),
      italic_(false),
      superscript_(false),
      subscript_(false),
      underline_(underline_style::none),
      strikethrough_(false)
{
}

font &font::bold(bool bold)
{
    bold_ = bold;
}

bool font::bold() const
{
    return bold_;
}

font &font::italic(bool italic)
{
    italic_ = italic;
}

bool font::italic() const
{
    return italic_;
}

font &font::strikethrough(bool strikethrough)
{
    strikethrough_ = strikethrough;
}

bool font::strikethrough() const
{
    return strikethrough_;
}

font &font::underline(underline_style new_underline)
{
    underline_ = new_underline;
}

bool font::underlined() const
{
    return underline_ != underline_style::none;
}

font::underline_style font::underline() const
{
    return underline_;
}

font &font::size(std::size_t size)
{
    size_ = size;
}

std::size_t font::size() const
{
    return size_;
}

font &font::name(const std::string &name)
{
    name_ = name;
}
std::string font::name() const
{
    return name_;
}

font &font::color(const xlnt::color &c)
{
    color_ = c;
}

font &font::family(std::size_t family)
{
    family_ = family;
}

font &font::scheme(const std::string &scheme)
{
    scheme_ = scheme;
}

optional<color> font::color() const
{
    return color_;
}

optional<std::size_t> font::family() const
{
    return family_;
}

optional<std::string> font::scheme() const
{
    return scheme_;
}

std::string font::to_hash_string() const
{
    std::string hash_string = "font";

    hash_string.append(std::to_string(bold_));
    hash_string.append(std::to_string(italic_));
    hash_string.append(std::to_string(superscript_));
    hash_string.append(std::to_string(subscript_));
    hash_string.append(std::to_string(strikethrough_));
    hash_string.append(name_);
    hash_string.append(std::to_string(size_));
    hash_string.append(std::to_string(static_cast<std::size_t>(underline_)));
    hash_string.append(family_ ? std::to_string(*family_) : "");
	hash_string.append(scheme_ ? *scheme_ : "");

    return hash_string;
}

} // namespace xlnt
