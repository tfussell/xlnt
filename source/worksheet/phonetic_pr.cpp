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
const std::array<std::string, 4> Types{
    "fullwidthKatakana",
    "halfwidthKatakana",
    "Hiragana",
    "noConversion"};

// Order of elements defined by phonetic_pr::Alignment enum
const std::array<std::string, 4> Alignments{
    "Center",
    "Distributed",
    "Left",
    "NoControl"};

} // namespace

namespace xlnt {
/// <summary>
/// out of line initialiser for static const member
/// </summary>
const std::string phonetic_pr::Serialised_ID = "phoneticPr";

phonetic_pr::phonetic_pr(Font_ID font)
    : font_id_(font)
{
}

void phonetic_pr::serialise(std::ostream &output_stream) const
{
    output_stream << '<' << Serialised_ID << R"( fontID=")" << std::to_string(font_id_) << '"';
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

phonetic_pr::Font_ID phonetic_pr::font_id() const
{
    return font_id_;
}

void phonetic_pr::font_id(Font_ID font)
{
    font_id_ = font;
}

bool phonetic_pr::has_type() const
{
    return type_.is_set();
}

phonetic_pr::Type phonetic_pr::type() const
{
    return type_.get();
}

void phonetic_pr::type(Type type)
{
    type_.set(type);
}

bool phonetic_pr::has_alignment() const
{
    return alignment_.is_set();
}

phonetic_pr::Alignment phonetic_pr::alignment() const
{
    return alignment_.get();
}

void phonetic_pr::alignment(Alignment align)
{
    alignment_.set(align);
}

// serialisation
const std::string &phonetic_pr::type_as_string(phonetic_pr::Type type)
{
    return Types[static_cast<int>(type)];
}

phonetic_pr::Type phonetic_pr::type_from_string(const std::string &str)
{
    for (int i = 0; i < Types.size(); ++i)
    {
        if (str == Types[i])
        {
            return static_cast<Type>(i);
        }
    }
    return Type::No_Conversion;
}

const std::string &phonetic_pr::alignment_as_string(Alignment type)
{
    return Alignments[static_cast<int>(type)];
}

phonetic_pr::Alignment phonetic_pr::alignment_from_string(const std::string &str)
{
    for (int i = 0; i < Alignments.size(); ++i)
    {
        if (str == Alignments[i])
        {
            return static_cast<Alignment>(i);
        }
    }
    return Alignment::No_Control;
}

} // namespace xlnt