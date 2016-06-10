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

#include <string>
#include <unordered_map>
#include <vector>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

class alignment;
class border;
class color;
class conditional_format;
class fill;
class font;
class format;
class base_format;
class style;
class number_format;
class protection;
class side;
class style;
class xml_document;
class xml_node;

namespace detail { struct stylesheet; }

/// <summary>
/// Reads and writes xl/styles.xml from/to an associated stylesheet.
/// </summary>
class XLNT_CLASS style_serializer
{
public:
    /// <summary>
    /// Construct a style_serializer which can write styles.xml based on wb or populate wb
    /// with styles from an existing styles.xml.
    style_serializer(detail::stylesheet &s);

    /// <summary>
    /// Load all styles from xml_document into workbook given in constructor.
    /// </summary>
    bool read_stylesheet(const xml_document &xml);

    /// <summary>
    /// Populate parameter xml with an XML tree representing the styles contained in the workbook
    /// given in the constructor.
    /// </summary>
    bool write_stylesheet(xml_document &xml);

private:
    /// <summary>
    /// Set in the constructor, this stylesheet is used as the source or target for all writing or reading, respectively.
    /// </summary>
    detail::stylesheet &stylesheet_;
};

} // namespace xlnt
