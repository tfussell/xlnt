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

#include <xlnt/xlnt_config.hpp>
#include <xlnt/styles/color.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

class style;

/// <summary>
/// Describes the font style of a particular cell.
/// </summary>
class XLNT_CLASS font : public hashable
{
public:
    enum class underline_style
    {
        none,
        double_,
        double_accounting,
        single,
        single_accounting
    };
    
    font();

    font &bold(bool bold);

    bool bold() const;

	font &superscript(bool superscript);

	bool superscript() const;

    font &italic(bool italic);

    bool italic() const;

    font &strikethrough(bool strikethrough);

    bool strikethrough() const;

    font &underline(underline_style new_underline);

    bool underlined() const;

    underline_style underline() const;

    font &size(std::size_t size);

    std::size_t size() const;

    font &name(const std::string &name);

    std::string name() const;

    font &color(const color &c);

	optional<xlnt::color> color() const;

    font &family(std::size_t family);

	optional<std::size_t> family() const;

    font &scheme(const std::string &scheme);

    optional<std::string> scheme() const;

protected:
    std::string to_hash_string() const override;

private:
    friend class style;

    std::string name_ = "Calibri";
    std::size_t size_ = 12;
    bool bold_ = false;
    bool italic_ = false;
    bool superscript_ = false;
    bool subscript_ = false;
    bool strikethrough_ = false;
	underline_style underline_ = underline_style::none;
	optional<xlnt::color> color_;
	optional<std::size_t> family_;
	optional<std::string> scheme_;
};

} // namespace xlnt
