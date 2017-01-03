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
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/workbook/named_range.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace {

/// <summary>
/// Return a vector containing string split at each delim.
/// </summary>
/// <remark>
/// This should maybe be in a utility header so it can be used elsewhere.
/// </remarks>
/*
std::vector<std::string> split_string(const std::string &string, char delim)
{
    std::vector<std::string> split;
    std::string::size_type previous_index = 0;
    auto separator_index = string.find(delim);

    while (separator_index != std::string::npos)
    {
        auto part = string.substr(previous_index, separator_index - previous_index);
        split.push_back(part);

        previous_index = separator_index + 1;
        separator_index = string.find(delim, previous_index);
    }

    split.push_back(string.substr(previous_index));

    return split;
}
*/

/*
std::vector<std::pair<std::string, std::string>> split_named_range(const std::string &named_range_string)
{
    std::vector<std::pair<std::string, std::string>> result;

    for (auto part : split_string(named_range_string, ','))
    {
        auto split = split_string(part, '!');

        if (split[0].front() == '\'' && split[0].back() == '\'')
        {
            split[0] = split[0].substr(1, split[0].length() - 2);
        }

        // Try to parse it. Use empty string if it's not a valid range.
        try
        {
            xlnt::range_reference ref(split[1]);
        }
        catch (xlnt::invalid_cell_reference)
        {
            split[1] = "";
        }

        result.push_back({ split[0], split[1] });
    }

    return result;
}
*/

} // namespace

namespace xlnt {

std::string named_range::name() const
{
    return name_;
}

const std::vector<named_range::target> &named_range::targets() const
{
    return targets_;
}

named_range::named_range()
{
}

named_range::named_range(const named_range &other)
{
    *this = other;
}

named_range::named_range(const std::string &name, const std::vector<named_range::target> &targets)
    : name_(name), targets_(targets)
{
}

named_range &named_range::operator=(const named_range &other)
{
    name_ = other.name_;
    targets_ = other.targets_;

    return *this;
}

} // namespace xlnt
