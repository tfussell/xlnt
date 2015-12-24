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

#include <cstdint>
#include <string>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/worksheet/footer.hpp>
#include <xlnt/worksheet/header.hpp>

namespace xlnt {

/// <summary>
/// Worksheet header and footer for all six possible positions.
/// </summary>
class XLNT_CLASS header_footer
{
  public:
    header_footer();

    header &get_left_header();
    header &get_center_header();
    header &get_right_header();
    
    footer &get_left_footer();
    footer &get_center_footer();
    footer &get_right_footer();

    bool is_default_header() const;
    bool is_default_footer() const;
    bool is_default() const;

  private:
    header left_header_, right_header_, center_header_;
    footer left_footer_, right_footer_, center_footer_;
};

} // namespace xlnt
