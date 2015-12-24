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

class range_reference;
class worksheet;

//TODO: why is this not in a class?
std::vector<std::pair<std::string, std::string>> XLNT_FUNCTION split_named_range(const std::string &named_range_string);

/// <summary>
/// A 2D range of cells in a worksheet that is referred to by name.
/// ws->get_range("A1:B2") could be replaced by ws->get_range("range1")
/// </summary>
class XLNT_CLASS named_range
{
  public:
    using target = std::pair<worksheet, range_reference>;

    named_range();
    named_range(const named_range &other);
    named_range(const std::string &name, const std::vector<target> &targets);

    std::string get_name() const;
    const std::vector<target> &get_targets() const;

    named_range &operator=(const named_range &other);

  private:
    std::string name_;
    std::vector<target> targets_;
};

} // namespace xlnt
