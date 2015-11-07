// Copyright (c) 2015 Thomas Fussell
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

#include <xlnt/utils/hash_combine.hpp>
#include <xlnt/styles/color.hpp>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

class style;

class XLNT_CLASS font
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

    void set_bold(bool bold)
    {
        bold_ = bold;
    }
    
    bool is_bold() const
    {
        return bold_;
    }

    void set_italic(bool italic)
    {
        italic_ = italic;
    }
    
    bool is_italic() const
    {
        return italic_;
    }

    void set_strikethrough(bool strikethrough)
    {
        strikethrough_ = strikethrough;
    }
    
    bool is_strikethrough() const
    {
        return strikethrough_;
    }

    void set_underline(underline_style new_underline)
    {
        underline_ = new_underline;
    }
    
    bool is_underline() const
    {
        return underline_ != underline_style::none;
    }
    
    underline_style get_underline() const
    {
        return underline_;
    }

    void set_size(std::size_t size)
    {
        size_ = size;
    }
    
    std::size_t get_size() const
    {
        return size_;
    }

    void set_name(const string &name)
    {
        name_ = name;
    }
    string get_name() const
    {
        return name_;
    }

    void set_color(color c)
    {
        color_ = c;
    }
    
    void set_family(std::size_t family)
    {
        has_family_ = true;
        family_ = family;
    }
    
    void set_scheme(const string &scheme)
    {
        has_scheme_ = true;
        scheme_ = scheme;
    }

    color get_color() const
    {
        return color_;
    }

    bool has_family() const
    {
        return has_family_;
    }
    
    std::size_t get_family() const
    {
        return family_;
    }

    bool has_scheme() const
    {
        return has_scheme_;
    }

    std::size_t hash() const
    {
        std::size_t seed = 0;

		hash_combine(seed, bold_);
		hash_combine(seed, italic_);
		hash_combine(seed, superscript_);
		hash_combine(seed, subscript_);
		hash_combine(seed, strikethrough_);
        hash_combine(seed, name_);
        hash_combine(seed, size_);
        hash_combine(seed, static_cast<std::size_t>(underline_));
		hash_combine(seed, color_.hash());
        hash_combine(seed, family_);
        hash_combine(seed, scheme_);

        return seed;
    }

    bool operator==(const font &other) const
    {
        return hash() == other.hash();
    }

  private:
    friend class style;

    string name_ = "Calibri";
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
    string scheme_;
};

} // namespace xlnt
