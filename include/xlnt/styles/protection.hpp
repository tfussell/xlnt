// Copyright (c) 2014-2021 Thomas Fussell
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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
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
    /// Returns an unlocked and unhidden protection object.
    /// </summary>
    static protection unlocked_and_visible();

    /// <summary>
    /// Returns a locked and unhidden protection object.
    /// </summary>
    static protection locked_and_visible();

    /// <summary>
    /// Returns an unlocked and hidden protection object.
    /// </summary>
    static protection unlocked_and_hidden();

    /// <summary>
    /// Returns a locked and hidden protection object.
    /// </summary>
    static protection locked_and_hidden();

    /// <summary>
    /// Constructs a default unlocked unhidden protection object.
    /// </summary>
    protection();

    /// <summary>
    /// Returns true if cells using this protection should be locked.
    /// </summary>
    bool locked() const;

    /// <summary>
    /// Sets the locked state of the protection to locked and returns a reference to the protection.
    /// </summary>
    protection &locked(bool locked);

    /// <summary>
    /// Returns true if cells using this protection should be hidden.
    /// </summary>
    bool hidden() const;

    /// <summary>
    /// Sets the hidden state of the protection to hidden and returns a reference to the protection.
    /// </summary>
    protection &hidden(bool hidden);

    /// <summary>
    /// Returns true if this protections is equivalent to right.
    /// </summary>
    bool operator==(const protection &other) const;

    /// <summary>
    /// Returns true if this protection is not equivalent to right.
    /// </summary>
    bool operator!=(const protection &other) const;

private:
    /// <summary>
    /// Whether the cell using this protection is locked or not
    /// </summary>
    bool locked_;

    /// <summary>
    /// Whether the cell using this protection is hidden or not
    /// </summary>
    bool hidden_;
};

} // namespace xlnt
