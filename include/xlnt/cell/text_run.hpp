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
	std::string get_string() const;

    /// <summary>
    ///
    /// </summary>
	void set_string(const std::string &string);

    /// <summary>
    ///
    /// </summary>
    bool has_size() const;

    /// <summary>
    ///
    /// </summary>
    std::size_t get_size() const;

    /// <summary>
    ///
    /// </summary>
    void set_size(std::size_t size);

    /// <summary>
    ///
    /// </summary>
    bool has_color() const;

    /// <summary>
    ///
    /// </summary>
    color get_color() const;

    /// <summary>
    ///
    /// </summary>
    void set_color(const color &new_color);

    /// <summary>
    ///
    /// </summary>
    bool has_font() const;

    /// <summary>
    ///
    /// </summary>
    std::string get_font() const;

    /// <summary>
    ///
    /// </summary>
    void set_font(const std::string &font);

    /// <summary>
    ///
    /// </summary>
    bool has_family() const;

    /// <summary>
    ///
    /// </summary>
    std::size_t get_family() const;

    /// <summary>
    ///
    /// </summary>
    void set_family(std::size_t family);

    /// <summary>
    ///
    /// </summary>
    bool has_scheme() const;

    /// <summary>
    ///
    /// </summary>
    std::string get_scheme() const;

    /// <summary>
    ///
    /// </summary>
    void set_scheme(const std::string &scheme);

    /// <summary>
    ///
    /// </summary>
    bool bold_set() const;

    /// <summary>
    ///
    /// </summary>
    bool is_bold() const;

    /// <summary>
    ///
    /// </summary>
    void set_bold(bool bold);

private:
    /// <summary>
    ///
    /// </summary>
	std::string string_;

    /// <summary>
    ///
    /// </summary>
    optional<std::size_t> size_;

    /// <summary>
    ///
    /// </summary>
    optional<color> color_;

    /// <summary>
    ///
    /// </summary>
    optional<std::string> font_;

    /// <summary>
    ///
    /// </summary>
    optional<std::size_t> family_;

    /// <summary>
    ///
    /// </summary>
    optional<std::string> scheme_;

    /// <summary>
    ///
    /// </summary>
    optional<bool> bold_;
};

} // namespace xlnt
