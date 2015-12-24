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
#include <vector>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

class cell_reference;
class tokenizer;

class XLNT_CLASS translator
{
    translator(const std::string &formula, const cell_reference &ref);

    std::vector<std::string> get_tokens();

    static std::string translate_row(const std::string &row_str, int row_delta);
    static std::string translate_col(const std::string &col_str, col_delta);

    std::pair<std::string, std::string> strip_ws_name(const std::string &range_str);

    void translate_range(const range_reference &range_ref);
    void translate_formula(const cell_reference &dest);

  private:
    const std::string ROW_RANGE_RE;
    const std::string COL_RANGE_RE;
    const std::string CELL_REF_RE;

    std::string formula_;
    cell_reference reference_;
    tokenizer tokenizer_;
};

} // namespace xlnt
