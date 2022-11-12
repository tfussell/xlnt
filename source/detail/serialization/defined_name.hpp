// Copyright (c) 2014-2021 Thomas Fussell
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

#include <string>

namespace xlnt {
namespace detail {

struct defined_name
{

    defined_name &operator=(const defined_name &other)
    {
        name = other.name;
        comment = other.comment;
        custom_menu = other.custom_menu;
        description = other.description;
        help = other.help;

        status_bar = other.status_bar;
        sheet_id = other.sheet_id;
        hidden = other.hidden;
        function = other.function;
        function_group_id = other.function_group_id;

        shortcut_key = other.shortcut_key;
        value = other.value;

        return *this;
    }

    bool operator==(const defined_name &other)
    {
        return name == other.name
            && comment == other.comment
            && custom_menu == other.custom_menu
            && description == other.description
            && help == other.help
            && status_bar == other.status_bar
            && sheet_id == other.sheet_id
            && hidden == other.hidden
            && function == other.function
            && function_group_id == other.function_group_id
            && shortcut_key == other.shortcut_key
            && value == other.value;
    }
    bool operator!=(const defined_name &other)
    {
        return !(*this == other);
    }

    // Implements (most of) CT_RevisionDefinedName, there's several "old" members in the spec but they're also ignored in other librarie
    std::string name;
    optional<std::string> comment;
    optional<std::string> custom_menu;
    optional<std::string> description;
    optional<std::string> help;
    optional<std::string> status_bar;
    optional<std::size_t> sheet_id; // 0 indexed.
    optional<bool> hidden;
    optional<bool> function;
    optional<std::size_t> function_group_id;
    optional<std::string> shortcut_key; // This is unsigned byte in the spec, but openpyxl uses string so let's try that
    std::string value;

    
};

} // namespace detail
} // namespace xlnt
