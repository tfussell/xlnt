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
#include <vector>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

class range_reference;
class worksheet;

/// <summary>
/// A 2D range of cells in a worksheet that is referred to by name.
/// ws->range("A1:B2") could be replaced by ws->range("range1")
/// </summary>
class XLNT_API named_range
{
public:
    /// <summary>
    /// Type alias for the combination of sheet and range this named_range points to.
    /// </summary>
    using target = std::pair<worksheet, range_reference>;

    /// <summary>
    /// Constructs a named range that has no name and has no targets.
    /// </summary>
    named_range();

    /// <summary>
    /// Constructs a named range by copying its name and targets from other.
    /// </summary>
    named_range(const named_range &other);

    /// <summary>
    /// Constructs a named range with the given name and targets.
    /// </summary>
    named_range(const std::string &name, const std::vector<target> &targets);

    /// <summary>
    /// Returns the name of this range.
    /// </summary>
    std::string name() const;

    /// <summary>
    /// Returns the set of targets of this named range as a vector.
    /// </summary>
    const std::vector<target> &targets() const;

    /// <summary>
    /// Assigns the name and targets of this named_range to that of other.
    /// </summary>
    named_range &operator=(const named_range &other);

    bool operator==(const named_range &rhs) const;

private:
    /// <summary>
    /// The name of this named range.
    /// </summary>
    std::string name_;

    /// <summary>
    /// The targets of this named range.
    /// </summary>
    std::vector<target> targets_;
};

} // namespace xlnt
