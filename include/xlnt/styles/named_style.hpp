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

#include <xlnt/styles/common_style.hpp>

namespace xlnt {

class workbook;

/// <summary>
/// Describes a style which has a name and can be applied to multiple individual
/// cell_styles.
/// </summary>
class XLNT_CLASS named_style : public common_style
{
public:
    named_style();
    named_style(const named_style &other);
    named_style &operator=(const named_style &other);

    std::string get_name() const;
    void set_name(const std::string &name);
    
    bool get_hidden() const;
    void set_hidden(bool value);
    
    std::size_t get_builtin_id() const;
    void set_builtin_id(std::size_t builtin_id);

protected:
    std::string to_hash_string() const override;

private:
    std::string name_;
    bool hidden_;
    std::size_t builtin_id;
};

} // namespace xlnt
