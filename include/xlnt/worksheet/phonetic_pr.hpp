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

#pragma once

#include <cstdint>
#include <ostream>
#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

/// <summary>
/// Phonetic properties
/// Element provides a collection of properties that affect display of East Asian Languages
/// [Serialised phoneticPr]
/// </summary>
class XLNT_API phonetic_pr
{
public:
    static const std::string Serialised_ID;

    /// <summary>
    /// possible values for alignment property
    /// </summary>
    enum class Alignment
    {
        Center,
        Distributed,
        Left,
        No_Control
    };

    /// <summary>
    /// possible values for type property
    /// </summary>
    enum class Type
    {
        Full_Width_Katakana,
        Half_Width_Katakana,
        Hiragana,
        No_Conversion
    };

    /// <summary>
    /// FontID represented by an unsigned 32-bit integer
    /// </summary>
    using Font_ID = std::uint32_t;

    /// <summary>
    /// Default ctor for phonetic properties
    /// </summary>
    phonetic_pr() = default;

    /// <summary>
    /// FontID ctor for phonetic properties
    /// </summary>
    explicit phonetic_pr(Font_ID font);

    /// <summary>
    /// adds the xml serialised representation of this element to the stream
    /// </summary>
    void serialise(std::ostream& output_stream) const;

    /// <summary>
    /// get the font index
    /// </summary>
    Font_ID font_id() const;

    /// <summary>
    /// set the font index
    /// </summary>
    void font_id(Font_ID font);

    /// <summary>
    /// is the phonetic type set
    /// </summary>
    bool has_type() const;

    /// <summary>
    /// returns the phonetic type
    /// </summary>
    Type type() const;

    /// <summary>
    /// sets the phonetic type
    /// </summary>
    void type(Type type);

    /// <summary>
    /// is the alignment set
    /// </summary>
    bool has_alignment() const;

    /// <summary>
    /// get the alignment
    /// </summary>
    Alignment alignment() const;

    /// <summary>
    /// set the alignment
    /// </summary>
    void alignment(Alignment align);

    // serialisation
    /// <summary>
    /// string form of the type enum
    /// </summary>
    static const std::string &type_as_string(Type type);

    /// <summary>
    /// type enum from string
    /// </summary>
    static Type type_from_string(const std::string &str);

    /// <summary>
    /// string form of alignment enum
    /// </summary>
    static const std::string &alignment_as_string(xlnt::phonetic_pr::Alignment type);

    /// <summary>
    /// alignment enum from string
    /// </summary>
    static Alignment alignment_from_string(const std::string &str);

private:
    /// <summary>
    /// zero based index into style sheet font record.
    /// Default: 0
    /// </summary>
    Font_ID font_id_ = 0;

    /// <summary>
    /// Type of characters to use.
    /// Default: full width katakana
    /// </summary>
    xlnt::optional<Type> type_;

    /// <summary>
    /// Alignment across the cell(s).
    /// Default: Left
    /// </summary>
    xlnt::optional<Alignment> alignment_;
};
}