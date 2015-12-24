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
#pragma once

#include <string>
#include <vector>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/packaging/default_type.hpp>
#include <xlnt/packaging/override_type.hpp>

namespace xlnt {

/// <summary>
/// The manifest keeps track of all files the OOXML package and
/// their type.
/// </summary>
class XLNT_CLASS manifest
{
  public:
    bool has_default_type(const std::string &extension) const;
    std::string get_default_type(const std::string &extension) const;
    const std::vector<default_type> &get_default_types() const;
    void add_default_type(const std::string &extension, const std::string &content_type);
    
    bool has_override_type(const std::string &part_name) const;
    std::string get_override_type(const std::string &part_name) const;
    const std::vector<override_type> &get_override_types() const;
    void add_override_type(const std::string &part_name, const std::string &content_type);

  private:
    std::vector<default_type> default_types_;
    std::vector<override_type> override_types_;
};

} // namespace xlnt
