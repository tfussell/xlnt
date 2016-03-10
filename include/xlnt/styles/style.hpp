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

#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/hashable.hpp>

namespace xlnt {

class workbook;

/// <summary>
/// Describes the entirety of the styling of a particular cell.
/// </summary>
class XLNT_CLASS style : public hashable
{
public:
    style();
    style(const style &other);
    style &operator=(const style &other);

    std::size_t hash() const;
    
    std::string get_name() const;
    void set_name(const std::string &name);
    
    std::size_t get_format_id() const;
    void set_format_id(std::size_t format_id);
    
    std::size_t get_builtin_id() const;
    void set_builtin_id(std::size_t builtin_id);
    
    void set_hidden(bool hidden);
    bool get_hidden() const;
    
protected:
    std::string to_hash_string() const override;
    
private:
    friend class workbook;
    std::size_t id_;
    std::string name_;
    std::size_t format_id_;
    std::size_t builtin_id_;
    bool hidden_;
};

} // namespace xlnt
