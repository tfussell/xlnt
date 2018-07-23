// Copyright (c) 2014-2018 Thomas Fussell
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

#include <cmath>
#include <cstdlib>

#include <xlnt/styles/color.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace {

std::array<std::uint8_t, 4> decode_hex_string(const std::string &hex_string)
{
    auto x = std::strtoul(hex_string.c_str(), nullptr, 16);

    auto a = static_cast<std::uint8_t>(x >> 24);
    auto r = static_cast<std::uint8_t>((x >> 16) & 0xff);
    auto g = static_cast<std::uint8_t>((x >> 8) & 0xff);
    auto b = static_cast<std::uint8_t>(x & 0xff);

    return { { r, g, b, a } };
}

} // namespace

namespace xlnt {

// indexed_color implementation

indexed_color::indexed_color(std::size_t index) : index_(index)
{
}

std::size_t indexed_color::index() const
{
    return index_;
}

void indexed_color::index(std::size_t index)
{
    index_ = index;
}

// theme_color implementation

theme_color::theme_color(std::size_t index) : index_(index)
{
}

std::size_t theme_color::index() const
{
    return index_;
}

void theme_color::index(std::size_t index)
{
    index_ = index;
}

// rgb_color implementation

std::string rgb_color::hex_string() const
{
    static const char *digits = "0123456789ABCDEF";
    std::string hex_string(8, '0');
    auto out_iter = hex_string.begin();

    for (auto byte : {rgba_[3], rgba_[0], rgba_[1], rgba_[2]})
    {
        for (auto i = 0; i < 2; ++i)
        {
            auto nibble = byte >> (4 * (1 - i)) & 0xf;
            *(out_iter++) = digits[nibble];
        }
    }

    return hex_string;
}

rgb_color::rgb_color(const std::string &hex_string)
    : rgba_(decode_hex_string(hex_string))
{
}

rgb_color::rgb_color(std::uint8_t r, std::uint8_t g, std::uint8_t b, std::uint8_t a)
    : rgba_({{r, g, b, a}})
{
}

std::uint8_t rgb_color::red() const
{
    return rgba_[0];
}

std::uint8_t rgb_color::green() const
{
    return rgba_[1];
}

std::uint8_t rgb_color::blue() const
{
    return rgba_[2];
}

std::uint8_t rgb_color::alpha() const
{
    return rgba_[3];
}

std::array<std::uint8_t, 3> rgb_color::rgb() const
{
    return {{red(), green(), blue()}};
}

std::array<std::uint8_t, 4> rgb_color::rgba() const
{
    return rgba_;
}

// color implementation

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

color::color() : color(indexed_color(0))
{
}

color::color(const rgb_color &rgb)
  : type_(color_type::rgb),
    rgb_(rgb),
    indexed_(0),
    theme_(0)
{
}

color::color(const indexed_color &indexed)
  : type_(color_type::indexed),
    rgb_(rgb_color(0, 0, 0, 0)),
    indexed_(indexed),
    theme_(0)
{
}

color::color(const theme_color &theme)
  : type_(color_type::theme),
    rgb_(rgb_color(0, 0, 0, 0)),
    indexed_(0),
    theme_(theme)
{
}

color_type color::type() const
{
    return type_;
}

bool color::auto_() const
{
    return auto__;
}

void color::auto_(bool value)
{
    auto__ = value;
}

const indexed_color& color::indexed() const
{
    assert_type(color_type::indexed);
    return indexed_;
}

indexed_color &color::indexed()
{
    assert_type(color_type::indexed);
    return indexed_;
}

const theme_color& color::theme() const
{
    assert_type(color_type::theme);
    return theme_;
}

theme_color &color::theme()
{
    assert_type(color_type::theme);
    return theme_;
}

const rgb_color& color::rgb() const
{
    assert_type(color_type::rgb);
    return rgb_;
}

rgb_color &color::rgb()
{
    assert_type(color_type::rgb);
    return rgb_;
}

void color::tint(double tint)
{
    tint_ = tint;
}

double color::tint() const
{
    return tint_;
}

void color::assert_type(color_type t) const
{
    if (t != type_)
    {
        throw invalid_attribute();
    }
}

bool color::operator==(const xlnt::color &other) const
{
    if (type_ != other.type_ || std::fabs(tint_ - other.tint_) != 0.0 || auto__ != other.auto__)
    {
        return false;
    }

    switch (type_)
    {
    case color_type::indexed:
        return indexed_.index() == other.indexed_.index();
    case color_type::theme:
        return theme_.index() == other.theme_.index();
    case color_type::rgb:
        return rgb_.hex_string() == other.rgb_.hex_string();
    }

    return false;
}

bool color::operator!=(const color &other) const
{
    return !(*this == other);
}

} // namespace xlnt
