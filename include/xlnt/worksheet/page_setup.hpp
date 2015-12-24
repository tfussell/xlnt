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
#include <xlnt/worksheet/orientation.hpp>
#include <xlnt/worksheet/page_break.hpp>
#include <xlnt/worksheet/paper_size.hpp>
#include <xlnt/worksheet/sheet_state.hpp>

namespace xlnt {

/// <summary>
/// Describes how a worksheet will be converted into a page during printing.
/// </summary>
struct XLNT_CLASS page_setup
{
  public:
    page_setup();
    
    bool is_default() const;

    page_break get_break() const;

    void set_break(page_break b);

    sheet_state get_sheet_state() const;

    void set_sheet_state(sheet_state sheet_state);

    paper_size get_paper_size() const;

    void set_paper_size(paper_size paper_size);

    orientation get_orientation() const;

    void set_orientation(orientation orientation);

    bool fit_to_page() const;

    void set_fit_to_page(bool fit_to_page);

    bool fit_to_height() const;

    void set_fit_to_height(bool fit_to_height);

    bool fit_to_width() const;

    void set_fit_to_width(bool fit_to_width);

    void set_horizontal_centered(bool horizontal_centered);

    bool get_horizontal_centered() const;

    void set_vertical_centered(bool vertical_centered);

    bool get_vertical_centered() const;

    void set_scale(double scale);
    
    double get_scale() const;

  private:
    bool default_;
    page_break break_;
    sheet_state sheet_state_;
    paper_size paper_size_;
    orientation orientation_;
    bool fit_to_page_;
    bool fit_to_height_;
    bool fit_to_width_;
    bool horizontal_centered_;
    bool vertical_centered_;
    double scale_;
};

} // namespace xlnt
