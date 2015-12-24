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
#include <xlnt/utils/hash_combine.hpp>

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

    void set_bold(bool bold);
    
    bool is_bold() const;

    void set_italic(bool italic);
    
    bool is_italic() const;

    void set_strikethrough(bool strikethrough);
    
    bool is_strikethrough() const;

    void set_underline(underline_style new_underline);
    
    bool is_underline() const;
    
    underline_style get_underline() const;

    void set_size(std::size_t size);
    
    std::size_t get_size() const;

    void set_name(const std::string &name);
    
    std::string get_name() const;

    void set_color(color c);
    
    void set_family(std::size_t family);
    
    void set_scheme(const std::string &scheme);

    color get_color() const;

    bool has_family() const;
    
    std::size_t get_family() const;

    bool has_scheme() const;
    
protected:
    std::string to_hash_string() const override;

  private:
    friend class style;

    std::string name_ = "Calibri";
    std::size_t size_ = 11;
    bool bold_ = false;
    bool italic_ = false;
    bool superscript_ = false;
    bool subscript_ = false;
    underline_style underline_ = underline_style::none;
    bool strikethrough_ = false;
    color color_;
    bool has_family_ = false;
    std::size_t family_;
    bool has_scheme_ = false;
    std::string scheme_;
};

} // namespace xlnt
