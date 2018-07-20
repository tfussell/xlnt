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

/// <summary>
/// Properties governing how the data in a workbook should be protected.
/// These values can be ignored by consumers.
/// </summary>
class XLNT_API document_security
{
public:
    /// <summary>
    /// Holds data describing the verifier that locks revisions or a workbook.
    /// </summary>
    struct lock_verifier
    {
        /// <summary>
        /// The algorithm used to create and verify this lock.
        /// </summary>
        std::string hash_algorithm;

        /// <summary>
        /// The initial salt combined with the password used to prevent rainbow table attacks
        /// </summary>
        std::string salt;

        /// <summary>
        /// The actual hash value represented as a string
        /// </summary>
        std::string hash;

        /// <summary>
        /// The number of times the hash should be applied to the password combined with the salt.
        /// This allows the difficulty of the hash to be increased as computing power increases.
        /// </summary>
        std::size_t spin_count;
    };

    /// <summary>
    /// Constructs a new document security object with default values.
    /// </summary>
    document_security() = default;

    /// <summary>
    /// If true, the workbook is locked for revisions.
    /// </summary>
    bool lock_revision = false;

    /// <summary>
    /// If true, worksheets can't be moved, renamed, (un)hidden, inserted, or deleted.
    /// </summary>
    bool lock_structure = false;

    /// <summary>
    /// If true, workbook windows will be opened at the same position with the same size
    /// every time they are loaded.
    /// </summary>
    bool lock_windows = false;

    /// <summary>
    /// The settings to allow the revision lock to be removed.
    /// </summary>
    lock_verifier revision_lock;

    /// <summary>
    /// The settings to allow the structure and windows lock to be removed.
    /// </summary>
    lock_verifier workbook_lock;
};

} // namespace xlnt
