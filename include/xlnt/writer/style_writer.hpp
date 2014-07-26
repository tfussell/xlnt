// Copyright (c) 2014 Thomas Fussell
// Copyright (c) 2010-2014 openpyxl
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
#include <vector>

#include "../styles/style.hpp"

namespace xlnt {

class workbook;

class style_writer
{
public:
    style_writer(workbook &wb);
    style_writer(const style_writer &);
    style_writer &operator=(const style_writer &);
    std::unordered_map<std::size_t, std::string> get_style_by_hash() const;
    std::string write_table() const;
    std::vector<style> get_styles() const;
    
private:
    std::vector<style> get_style_list(const workbook &wb) const;
    std::unordered_map<int, std::string> write_fonts() const;
    std::unordered_map<int, std::string> write_fills() const;
    std::unordered_map<int, std::string> write_borders() const;
    void write_cell_style_xfs();
    void write_cell_xfs();
    void write_cell_style();
    void write_dxfs();
    void write_table_styles();
    void write_number_formats();
    
    std::vector<style> style_list_;
    workbook &wb_;
};
    
} // namespace xlnt
