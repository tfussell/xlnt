// Copyright (c) 2018 Thomas Fussell
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
#include <xlnt/xlnt_config.hpp>

namespace xlnt {

namespace detail {
struct hyperlink_impl;
}

class cell;
class range;
class relationship;

/// <summary>
/// Describes a hyperlink pointing from a cell to another cell or a URL.
/// </summary>
class XLNT_API hyperlink
{
public:
    bool external() const;
    class relationship relationship() const;
    // external target
    std::string url() const;
    // internal target
    std::string target_range() const;

    bool has_display() const;
    void display(const std::string &value);
    const std::string& display() const;

    bool has_tooltip() const;
    void tooltip(const std::string &value);
    const std::string& tooltip() const;

    bool has_location() const;
    void location(const std::string &value);
    const std::string& location() const;

private:
    friend class cell;
    hyperlink(detail::hyperlink_impl *d);
    detail::hyperlink_impl *d_;
};

} // namespace xlnt

