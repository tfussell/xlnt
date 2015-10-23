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

#include <xlnt/common/hash_combine.hpp>
#include <xlnt/styles/color.hpp>

namespace xlnt {
    
class style;

class font
{
public:
    enum class underline
    {
        none,
        double_,
        double_accounting,
        single,
        single_accounting
    };
    
    void set_bold(bool bold) { bold_ = bold; }
    bool is_bold() const { return bold_; }
    
    void set_size(int size) { size_ = size; }
    int get_size() const { return size_; }
    
    void set_name(const std::string &name) { name_ = name; }
    std::string get_name() const { return name_; }
    
    void set_color(color c) { color_ = c; }
    void set_family(int family) { has_family_ = true; family_ = family; }
    void set_scheme(const std::string &scheme) { has_scheme_ = true; scheme_ = scheme; }
    
    color get_color() const { return color_; }
    
    bool has_family() const { return has_family_; }
    bool has_scheme() const { return has_scheme_; }
    
    std::size_t hash() const
    {
        std::size_t seed = bold_;
        seed = seed << 1 & italic_;
        seed = seed << 1 & superscript_;
        seed = seed << 1 & subscript_;
        seed = seed << 1 & strikethrough_;
        hash_combine(seed, name_);
        hash_combine(seed, size_);
        hash_combine(seed, static_cast<std::size_t>(underline_));
        hash_combine(seed, static_cast<std::size_t>(color_.get_type()));
        switch(color_.get_type())
        {
            case color::type::indexed: hash_combine(seed, color_.get_index()); break;
            case color::type::theme: hash_combine(seed, color_.get_theme()); break;
            default: break;
        }
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
    
     std::string name_ = "Calibri";
     int size_ = 11;
     bool bold_ = false;
     bool italic_ = false;
     bool superscript_ = false;
     bool subscript_ = false;
     underline underline_ = underline::none;
     bool strikethrough_ = false;
     color color_;
     bool has_family_ = false;
     int family_;
     bool has_scheme_ = false;
     std::string scheme_;
};

} // namespace xlnt
