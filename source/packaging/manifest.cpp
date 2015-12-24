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

namespace {

bool match_path(const std::string &path, const std::string &comparand)
{
    if (path == comparand)
    {
        return true;
    }

    if (path[0] != '/' && path[0] != '.' && comparand[0] == '/')
    {
        return match_path("/" + path, comparand);
    }
    else if (comparand[0] != '/' && comparand[0] != '.' && path[0] == '/')
    {
        return match_path(path, "/" + comparand);
    }

    return false;
}

} // namespace

namespace xlnt {

bool manifest::has_default_type(const std::string &extension) const
{
    return std::find_if(default_types_.begin(), default_types_.end(),
                        [&](const default_type &d) { return d.get_extension() == extension; }) != default_types_.end();
}

bool manifest::has_override_type(const std::string &part_name) const
{
    return std::find_if(override_types_.begin(), override_types_.end(), [&](const override_type &d) {
               return match_path(d.get_part_name(), part_name);
           }) != override_types_.end();
}

void manifest::add_default_type(const std::string &extension, const std::string &content_type)
{
    if(has_default_type(extension)) return;
    default_types_.push_back(default_type(extension, content_type));
}

void manifest::add_override_type(const std::string &part_name, const std::string &content_type)
{
    if(has_override_type(part_name)) return;
    override_types_.push_back(override_type(part_name, content_type));
}

std::string manifest::get_default_type(const std::string &extension) const
{
    auto match = std::find_if(default_types_.begin(), default_types_.end(),
                              [&](const default_type &d) { return d.get_extension() == extension; });

    if (match == default_types_.end())
    {
        throw std::runtime_error("no default type found for extension: " + extension);
    }

    return match->get_content_type();
}

std::string manifest::get_override_type(const std::string &part_name) const
{
    auto match = std::find_if(override_types_.begin(), override_types_.end(),
                              [&](const override_type &d) { return match_path(d.get_part_name(), part_name); });

    if (match == override_types_.end())
    {
        throw std::runtime_error("no default type found for part name: " + part_name);
    }

    return match->get_content_type();
}

const std::vector<default_type> &manifest::get_default_types() const
{
    return default_types_;
}

const std::vector<override_type> &manifest::get_override_types() const
{
    return override_types_;
}

} // namespace xlnt
