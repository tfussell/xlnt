// Copyright (c) 2016 Thomas Fussell
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

#include <xlnt/xlnt_config.hpp>
#include <xlnt/styles/color.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

//todo: should this just be a struct?

/// <summary>
/// A formatted string
/// </summary>
class XLNT_API text_run 
{
public:
    /// <summary>
    ///
    /// </summary>
	text_run();

    /// <summary>
    ///
    /// </summary>
	text_run(const std::string &string);

    /// <summary>
    ///
    /// </summary>
    bool has_formatting() const;
    
    /// <summary>
    ///
    /// </summary>
	std::string string() const;

    /// <summary>
    ///
    /// </summary>
	void string(const std::string &string);

    /// <summary>
    ///
    /// </summary>
    bool has_size() const;

    /// <summary>
    ///
    /// </summary>
    std::size_t size() const;

    /// <summary>
    ///
    /// </summary>
    void size(std::size_t size);

    /// <summary>
    ///
    /// </summary>
    bool has_color() const;

    /// <summary>
    ///
    /// </summary>
    color color() const;

    /// <summary>
    ///
    /// </summary>
    void color(const class color &new_color);

    /// <summary>
    ///
    /// </summary>
    bool has_font() const;

    /// <summary>
    ///
    /// </summary>
    std::string font() const;

    /// <summary>
    ///
    /// </summary>
    void font(const std::string &font);

    /// <summary>
    ///
    /// </summary>
    bool has_family() const;

    /// <summary>
    ///
    /// </summary>
    std::size_t family() const;

    /// <summary>
    ///
    /// </summary>
    void family(std::size_t family);

    /// <summary>
    ///
    /// </summary>
    bool has_scheme() const;

    /// <summary>
    ///
    /// </summary>
    std::string scheme() const;

    /// <summary>
    ///
    /// </summary>
    void scheme(const std::string &scheme);

    /// <summary>
    ///
    /// </summary>
    bool bold() const;

    /// <summary>
    ///
    /// </summary>
    void bold(bool bold);

    /// <summary>
    ///
    /// </summary>
    bool underline_set() const;

    /// <summary>
    ///
    /// </summary>
    bool underlined() const;

    /// <summary>
    ///
    /// </summary>
    font::underline_style underline() const;

    /// <summary>
    ///
    /// </summary>
    void underline(font::underline_style style);

private:
    /// <summary>
    ///
    /// </summary>
	std::string string_;

    /// <summary>
    ///
    /// </summary>
    optional<xlnt::font> font_;
};

} // namespace xlnt
