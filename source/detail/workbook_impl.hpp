// Copyright (c) 2014-2016 Thomas Fussell
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

#include <iterator>
#include <vector>

#include <detail/stylesheet.hpp>
#include <detail/worksheet_impl.hpp>
#include <xlnt/workbook/encoding.hpp>
#include <xlnt/packaging/app_properties.hpp>
#include <xlnt/packaging/document_properties.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/workbook/theme.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/sheet_view.hpp>

namespace xlnt {
namespace detail {

struct worksheet_impl;

struct workbook_impl
{
    workbook_impl();

    workbook_impl(const workbook_impl &other)
        : active_sheet_index_(other.active_sheet_index_),
          worksheets_(other.worksheets_),
          relationships_(other.relationships_),
          root_relationships_(other.root_relationships_),
          shared_strings_(other.shared_strings_),
          properties_(other.properties_),
          app_properties_(other.app_properties_),
          guess_types_(other.guess_types_),
          data_only_(other.data_only_),
          read_only_(other.read_only_),
          stylesheet_(other.stylesheet_),
          manifest_(other.manifest_)
    {
    }

    workbook_impl &operator=(const workbook_impl &other)
    {
        active_sheet_index_ = other.active_sheet_index_;
        worksheets_.clear();
        std::copy(other.worksheets_.begin(), other.worksheets_.end(), back_inserter(worksheets_));
        relationships_.clear();
        std::copy(other.relationships_.begin(), other.relationships_.end(), std::back_inserter(relationships_));
        root_relationships_.clear();
        std::copy(other.root_relationships_.begin(), other.root_relationships_.end(),
                  std::back_inserter(root_relationships_));
        shared_strings_.clear();
        std::copy(other.shared_strings_.begin(), other.shared_strings_.end(), std::back_inserter(shared_strings_));
        properties_ = other.properties_;
        app_properties_ = other.app_properties_;
        guess_types_ = other.guess_types_;
        data_only_ = other.data_only_;
        read_only_ = other.read_only_;
        manifest_ = other.manifest_;

        return *this;
    }

    std::size_t active_sheet_index_;
    std::vector<worksheet_impl> worksheets_;
    std::vector<relationship> relationships_;
    std::vector<relationship> root_relationships_;
    std::vector<text> shared_strings_;

    document_properties properties_;
    app_properties app_properties_;

    bool guess_types_;
    bool data_only_;
    bool read_only_;

    stylesheet stylesheet_;
    
    manifest manifest_;
    theme theme_;
    std::vector<std::uint8_t> thumbnail_;
};

} // namespace detail
} // namespace xlnt
