// Copyright (c) 2016 Thomas Fussell
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
#include <xlnt/cell/text_run.hpp>

namespace xlnt {

/// <summary>
/// Encapsulates zero or more formatted text runs where a text run
/// is a string of text with the same defined formatting.
/// </summary>
class XLNT_API formatted_text
{
public:
    /// <summary>
    /// Remove all text runs from this text.
    /// </summary>
    void clear();

    /// <summary>
    /// Clear any runs in this text and add a single run with default formatting and
    /// the given string as its textual content.
    /// </summary>
    void plain_text(const std::string &s);

    /// <summary>
    /// Combine the textual content of each text run in order and return the result.
    /// </summary>
    std::string plain_text() const;

    /// <summary>
    /// Returns a copy of the individual runs that comprise this text.
    /// </summary>
    std::vector<text_run> runs() const;

    /// <summary>
    /// Set the runs of this text all at once.
    /// </summary>
    void runs(const std::vector<text_run> &new_runs);

    /// <summary>
    /// Add a new run to the end of the set of runs.
    /// </summary>
    void add_run(const text_run &t);

    /// <summary>
    /// Returns true if the runs that make up this text are identical to those in rhs.
    /// </summary>
    bool operator==(const formatted_text &rhs) const;

private:
    /// <summary>
    ///
    /// </summary>
    std::vector<text_run> runs_;
};

} // namespace xlnt
