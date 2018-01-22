// Copyright (c) 2014-2018 Thomas Fussell
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

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

// TOOD: does this really need its own class?

/// <summary>
/// Protection applied to a particular worksheet to prevent it from being modified.
/// </summary>
class XLNT_API sheet_protection
{
public:
    /// <summary>
    /// Calculates and returns the hash of the given protection password.
    /// </summary>
    static std::string hash_password(const std::string &password);

    /// <summary>
    /// Sets the protection password to password.
    /// </summary>
    void password(const std::string &password);

    /// <summary>
    /// Returns the hash of the password set for this sheet protection.
    /// </summary>
    std::string hashed_password() const;

private:
    /// <summary>
    /// The hash of the password.
    /// </summary>
    std::string hashed_password_;
};

} // namespace xlnt
