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

#include <cstdlib>

#include <xlnt/styles/color.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace xlnt {

indexed_color::indexed_color(std::size_t index) : index_(index)
{
}

theme_color::theme_color(std::size_t index) : index_(index)
{
}

const color color::black()
{
    return color(rgb_color("ff000000"));
}

const color color::white()
{
    return color(rgb_color("ffffffff"));
}

const color color::red()
{
    return color(rgb_color("ffff0000"));
}

const color color::darkred()
{
    return color(rgb_color("ff8b0000"));
}

const color color::blue()
{
    return color(rgb_color("ff0000ff"));
}

const color color::darkblue()
{
    return color(rgb_color("ff00008b"));
}

const color color::green()
{
    return color(rgb_color("ff00ff00"));
}

const color color::darkgreen()
{
    return color(rgb_color("ff008b00"));
}

const color color::yellow()
{
    return color(rgb_color("ffffff00"));
}

const color color::darkyellow()
{
    return color(rgb_color("ffcccc00"));
}

color::color()
    : type_(type::auto_),
      rgb_(rgb_color(0, 0, 0, 0)),
      indexed_(0),
      theme_(0),
	  tint_(0)
{
}

color::color(const rgb_color &rgb)
	: type_(type::rgb),
	  rgb_(rgb),
	  indexed_(0),
	  theme_(0),
	  tint_(0)
{
}

color::color(const indexed_color &indexed)
	: type_(type::indexed),
      rgb_(rgb_color(0, 0, 0, 0)),
      indexed_(indexed),
      theme_(0),
	  tint_(0)
{
}

color::color(const theme_color &theme)
	: type_(type::theme),
      rgb_(rgb_color(0, 0, 0, 0)),
      indexed_(0),
      theme_(theme),
	  tint_(0)
{
}

color::type color::get_type() const
{
    return type_;
}

bool color::is_auto() const
{
	return type_ == type::auto_;
}

const indexed_color &color::get_indexed() const
{
	assert_type(type::indexed);
	return indexed_;
}

const theme_color &color::get_theme() const
{
	assert_type(type::theme);
	return theme_;
}

std::string rgb_color::get_hex_string() const
{
	static const char* digits = "0123456789abcdef";
	std::string hex_string(8, '0');
	auto out_iter = hex_string.begin();

	for (auto byte : { rgba_[3], rgba_[0], rgba_[1], rgba_[2] })
	{
		for (auto i = 0; i < 2; ++i)
		{
			auto nibble = byte >> (4 * (1 - i)) & 0xf;
			*(out_iter++) = digits[nibble];
		}
	}

	return hex_string;
}

std::array<std::uint8_t, 4> rgb_color::decode_hex_string(const std::string &hex_string)
{
	auto x = std::strtoul(hex_string.c_str(), NULL, 16);

	auto a = static_cast<std::uint8_t>(x >> 24);
	auto r = static_cast<std::uint8_t>((x >> 16) & 0xff);
	auto g = static_cast<std::uint8_t>((x >> 8) & 0xff);
	auto b = static_cast<std::uint8_t>(x & 0xff);

	return { {r, g, b, a} };
}

rgb_color::rgb_color(const std::string &hex_string)
	: rgba_(decode_hex_string(hex_string))
{
}

rgb_color::rgb_color(std::uint8_t r, std::uint8_t g, std::uint8_t b, std::uint8_t a)
	: rgba_({{r, g, b, a}})
{
}

std::size_t indexed_color::get_index() const
{
	return index_;
}

std::size_t theme_color::get_index() const
{
	return index_;
}

const rgb_color &color::get_rgb() const
{
	assert_type(type::rgb);
	return rgb_;
}

void color::set_tint(double tint)
{
	tint_ = tint;
}

void color::assert_type(type t) const
{
	if (t != type_)
	{
		throw invalid_attribute();
	}
}

bool color::operator==(const xlnt::color &other) const
{
    if (type_ != other.type_ || tint_ != other.tint_) return false;
    switch(type_)
    {
        case type::auto_:
        case type::indexed :
            return indexed_.get_index() == other.indexed_.get_index();
        case type::theme:
            return theme_.get_index() == other.theme_.get_index();
        case type::rgb:
            return rgb_.get_hex_string() == other.rgb_.get_hex_string();
    }
    return false;
}

} // namespace xlnt
