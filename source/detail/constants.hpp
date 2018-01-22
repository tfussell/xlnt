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

#include <string>
#include <unordered_map>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/index_types.hpp>
#include <xlnt/utils/path.hpp>

namespace xlnt {

struct XLNT_API constants
{
    /// <summary>
    /// Returns the lowest allowable row index in a worksheet.
    /// </summary>
    static row_t min_row();

    /// <summary>
    /// Returns the largest allowable row index in a worksheet.
    /// </summary>
    static row_t max_row();

    /// <summary>
    /// Returns the lowest allowable column index in a worksheet.
    /// </summary>
    static const column_t min_column();

    /// <summary>
    /// Returns the largest allowable column index in a worksheet.
    /// </summary>
    static const column_t max_column();

    /// <summary>
    /// Returns the URI of the directory containing package properties.
    /// </summary>
    static const path package_properties();

    /// <summary>
    /// Returns the URI of the directory containing SpreatsheetML package parts.
    /// </summary>
    static const path package_xl();

    /// <summary>
    /// Returns the URI of the directory containing root relationships package part.
    /// </summary>
    static const path package_root_rels();

    /// <summary>
    /// Returns the URI of the directory containing package themes.
    /// </summary>
    static const path package_theme();

    /// <summary>
    /// Returns the URI of the directory containing package worksheets.
    /// </summary>
    static const path package_worksheets();

	/// <summary>
	/// Returns the URI of the directory containing package worksheets.
	/// </summary>
	static const path package_drawings();

    /// <summary>
    /// Returns the URI of the content types package part.
    /// </summary>
    static const path part_content_types();

    /// <summary>
    /// Returns the URI of the core properties package part.
    /// </summary>
    static const path part_core();

    /// <summary>
    /// Returns the URI of the app properties package part.
    /// </summary>
    static const path part_app();

    /// <summary>
    /// Returns the URI of the workbook package part.
    /// </summary>
    static const path part_workbook();

    /// <summary>
    /// Returns the URI of the root relationships package part.
    /// </summary>
    static const path part_root_relationships();

    /// <summary>
    /// Returns the URI of the styles package part.
    /// </summary>
    static const path part_styles();

    /// <summary>
    /// Returns the URI of the theme package part.
    /// </summary>
    static const path part_theme();

    /// <summary>
    /// Returns the URI of the shared strings package part.
    /// </summary>
    static const path part_shared_strings();

    /// <summary>
    /// Returns an unordered_map mapping namespace names to namespaces.
    /// </summary>
    static const std::unordered_map<std::string, std::string> &namespaces();

    /// <summary>
    /// Returns the namespace URI from a namespace name.
    /// </summary>
    static const std::string &ns(const std::string &id);
};

} // namespace xlnt
