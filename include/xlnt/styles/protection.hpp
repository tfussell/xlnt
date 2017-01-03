// Copyright (c) 2014-2017 Thomas Fussell
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

#include <cstddef>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

/// <summary>
/// Describes the protection style of a particular cell.
/// </summary>
class XLNT_API protection
{
public:
    /// <summary>
    ///
    /// </summary>
    static protection unlocked_and_visible();

    /// <summary>
    ///
    /// </summary>
    static protection locked_and_visible();

    /// <summary>
    ///
    /// </summary>
    static protection unlocked_and_hidden();

    /// <summary>
    ///
    /// </summary>
    static protection locked_and_hidden();

    /// <summary>
    ///
    /// </summary>
    protection();

    /// <summary>
    ///
    /// </summary>
    bool locked() const;

    /// <summary>
    ///
    /// </summary>
    protection &locked(bool locked);

    /// <summary>
    ///
    /// </summary>
    bool hidden() const;

    /// <summary>
    ///
    /// </summary>
    protection &hidden(bool hidden);

    /// <summary>
    /// Returns true if left is exactly equal to right.
    /// </summary>
    XLNT_API friend bool operator==(const protection &left, const protection &right);

    /// <summary>
    /// Returns true if left is not exactly equal to right.
    /// </summary>
    XLNT_API friend bool operator!=(const protection &left, const protection &right)
    {
        return !(left == right);
    }

private:
    /// <summary>
    ///
    /// </summary>
    bool locked_;

    /// <summary>
    ///
    /// </summary>
    bool hidden_;
};

} // namespace xlnt
