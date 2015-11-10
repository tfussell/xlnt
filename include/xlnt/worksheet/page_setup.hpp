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

#include "xlnt_config.hpp"

namespace xlnt {

struct XLNT_CLASS page_setup
{
    enum class page_break
    {
        none = 0,
        row = 1,
        column = 2
    };

    enum class sheet_state
    {
        visible,
        hidden,
        very_hidden
    };

    enum class paper_size
    {
        letter = 1,
        letter_small = 2,
        tabloid = 3,
        ledger = 4,
        legal = 5,
        statement = 6,
        executive = 7,
        a3 = 8,
        a4 = 9,
        a4_small = 10,
        a5 = 11
    };

    enum class orientation
    {
        portrait,
        landscape
    };

  public:
    page_setup()
        : default_(true),
          break_(page_break::none),
          sheet_state_(sheet_state::visible),
          paper_size_(paper_size::letter),
          orientation_(orientation::portrait),
          fit_to_page_(false),
          fit_to_height_(false),
          fit_to_width_(false),
          horizontal_centered_(false),
          vertical_centered_(false),
          scale_(1)
    {
    }
    bool is_default() const
    {
        return default_;
    }
    page_break get_break() const
    {
        return break_;
    }
    void set_break(page_break b)
    {
        default_ = false;
        break_ = b;
    }
    sheet_state get_sheet_state() const
    {
        return sheet_state_;
    }
    void set_sheet_state(sheet_state sheet_state)
    {
        sheet_state_ = sheet_state;
    }
    paper_size get_paper_size() const
    {
        return paper_size_;
    }
    void set_paper_size(paper_size paper_size)
    {
        default_ = false;
        paper_size_ = paper_size;
    }
    orientation get_orientation() const
    {
        return orientation_;
    }
    void set_orientation(orientation orientation)
    {
        default_ = false;
        orientation_ = orientation;
    }
    bool fit_to_page() const
    {
        return fit_to_page_;
    }
    void set_fit_to_page(bool fit_to_page)
    {
        default_ = false;
        fit_to_page_ = fit_to_page;
    }
    bool fit_to_height() const
    {
        return fit_to_height_;
    }
    void set_fit_to_height(bool fit_to_height)
    {
        default_ = false;
        fit_to_height_ = fit_to_height;
    }
    bool fit_to_width() const
    {
        return fit_to_width_;
    }
    void set_fit_to_width(bool fit_to_width)
    {
        default_ = false;
        fit_to_width_ = fit_to_width;
    }
    void set_horizontal_centered(bool horizontal_centered)
    {
        default_ = false;
        horizontal_centered_ = horizontal_centered;
    }
    bool get_horizontal_centered() const
    {
        return horizontal_centered_;
    }
    void set_vertical_centered(bool vertical_centered)
    {
        default_ = false;
        vertical_centered_ = vertical_centered;
    }
    bool get_vertical_centered() const
    {
        return vertical_centered_;
    }
    void set_scale(double scale)
    {
        default_ = false;
        scale_ = scale;
    }
    double get_scale() const
    {
        return scale_;
    }

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
