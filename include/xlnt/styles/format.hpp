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

#include <xlnt/styles/formattable.hpp>

namespace xlnt {

class style;
namespace detail { struct workbook_impl; }

/// <summary>
/// Describes the formatting of a particular cell.
/// </summary>
class XLNT_CLASS format : public formattable
{
public:
    format();
    format(const format &other);
    format &operator=(const format &other);
    
    // Named Style
    bool has_style() const;
    std::string get_style_name() const;
    style &get_style();
    const style &get_style() const;
    void set_style(const std::string &name);
    void set_style(const style &new_style);
    void remove_style();
    
protected:
    std::string to_hash_string() const override;

private:
    detail::workbook_impl *parent_;
    std::string named_style_name_;
};

} // namespace xlnt
