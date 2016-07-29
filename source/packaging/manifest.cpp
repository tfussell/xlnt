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
#include <algorithm>
#include <stdexcept>

#include <xlnt/packaging/manifest.hpp>

namespace xlnt {

void manifest::clear()
{
    default_types_.clear();
    override_types_.clear();
}

bool manifest::has_default_type(const std::string &extension) const
{
    return default_types_.find(extension) != default_types_.end();
}

bool manifest::has_override_type(const std::string &part_name) const
{
    auto absolute = part_name.front() == '/' ? part_name : ("/" + part_name);
    return override_types_.find(absolute) != override_types_.end();
}

void manifest::add_default_type(const std::string &extension, const std::string &content_type)
{
    default_types_[extension] = default_type(extension, content_type);
}

void manifest::add_override_type(const std::string &part_name, const std::string &content_type)
{
    auto absolute = part_name.front() == '/' ? part_name : ("/" + part_name);
    override_types_[absolute] = override_type(absolute, content_type);
}

std::string manifest::get_default_type(const std::string &extension) const
{
    return default_types_.at(extension).get_content_type();
}

std::string manifest::get_override_type(const std::string &part_name) const
{
    auto absolute = part_name.front() == '/' ? part_name : ("/" + part_name);
    return override_types_.at(absolute).get_content_type();
}

const std::unordered_map<std::string, default_type> &manifest::get_default_types() const
{
    return default_types_;
}

const std::unordered_map<std::string, override_type> &manifest::get_override_types() const
{
    return override_types_;
}

} // namespace xlnt
