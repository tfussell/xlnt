// Copyright (c) 2014-2018 Thomas Fussell
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

#include <xlnt/worksheet/phonetic_pr.hpp>
#include <array>
namespace {
// Order of elements defined by phonetic_pr::Type enum
const std::array<std::string, 4> &Types()
{
    static const std::array<std::string, 4> types{
        std::string("fullwidthKatakana"),
        std::string("halfwidthKatakana"),
        std::string("Hiragana"),
        std::string("noConversion")};
    return types;
}

// Order of elements defined by phonetic_pr::alignment enum
const std::array<std::string, 4> &Alignments()
{
    static const std::array<std::string, 4> alignments{
        std::string("Center"),
        std::string("Distributed"),
        std::string("Left"),
        std::string("NoControl")};
    return alignments;
}

} // namespace

namespace xlnt {
/// <summary>
/// out of line initialiser for static const member
/// </summary>
std::string phonetic_pr::Serialised_ID()
{
    return "phoneticPr";
}

phonetic_pr::phonetic_pr(font_id_t font)
    : font_id_(font)
{
}

void phonetic_pr::serialise(std::ostream &output_stream) const
{
    output_stream << '<' << Serialised_ID() << R"( fontID=")" << std::to_string(font_id_) << '"';
    if (has_type())
    {
        output_stream << R"( type=")" << type_as_string(type_.get()) << '"';
    }
    if (has_alignment())
    {
        output_stream << R"( alignment=")" << alignment_as_string(alignment_.get()) << '"';
    }
    output_stream << "/>";
}

phonetic_pr::font_id_t phonetic_pr::font_id() const
{
    return font_id_;
}

void phonetic_pr::font_id(font_id_t font)
{
    font_id_ = font;
}

bool phonetic_pr::has_type() const
{
    return type_.is_set();
}

phonetic_pr::phonetic_type phonetic_pr::type() const
{
    return type_.get();
}

void phonetic_pr::type(phonetic_type type)
{
    type_.set(type);
}

bool phonetic_pr::has_alignment() const
{
    return alignment_.is_set();
}

phonetic_pr::align phonetic_pr::alignment() const
{
    return alignment_.get();
}

void phonetic_pr::alignment(align align)
{
    alignment_.set(align);
}

// serialisation
const std::string &phonetic_pr::type_as_string(phonetic_pr::phonetic_type type)
{
    return Types()[static_cast<std::size_t>(type)];
}

phonetic_pr::phonetic_type phonetic_pr::type_from_string(const std::string &str)
{
    for (std::size_t i = 0; i < Types().size(); ++i)
    {
        if (str == Types()[i])
        {
            return static_cast<phonetic_type>(i);
        }
    }
    return phonetic_type::no_conversion;
}

const std::string &phonetic_pr::alignment_as_string(align type)
{
    return Alignments()[static_cast<std::size_t>(type)];
}

phonetic_pr::align phonetic_pr::alignment_from_string(const std::string &str)
{
    for (std::size_t i = 0; i < Alignments().size(); ++i)
    {
        if (str == Alignments()[i])
        {
            return static_cast<align>(i);
        }
    }
    return align::no_control;
}

bool phonetic_pr::operator==(const phonetic_pr &rhs) const
{
    return font_id_ == rhs.font_id_
        && type_ == rhs.type_
        && alignment_ == rhs.alignment_;
}

} // namespace xlnt