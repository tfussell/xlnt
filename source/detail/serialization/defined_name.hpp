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

    bool operator==(const defined_name &other) const
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

    bool operator!=(const defined_name &other) const
    {
        return !(*this == other);
    }

    // Implements (most of) CT_RevisionDefinedName, there's several "old" members in the spec but they're also ignored in other librarie
    std::string name;                       // A string representing the name for this defined name.
    optional<std::string> comment;          // A string representing a comment about the defined name.
    optional<std::string> custom_menu;      // A string representing the new custom menu text
    optional<std::string> description;      // A string representing the new description text for the defined name.
    optional<std::string> help;             // A string representing the new help topic text.
    optional<std::string> status_bar;       // A string representing the new status bar text.
    optional<std::size_t> sheet_id;         // An integer representing the id of the sheet to which this defined name belongs. This shall be used local defined names only. 0 indexed indexed.
    optional<bool> hidden;                  // A Boolean value indicating whether the named range is now hidden.
    optional<bool> function;                // A Boolean value indicating that the defined name refers to a function. True if the defined name is a function, false otherwise.
    optional<std::size_t> function_group_id;// Represents the new function group id.
    optional<std::string> shortcut_key;     // Represents the new keyboard shortcut. This is unsigned byte in the spec, but openpyxl uses string so let's try that
    std::string value;                      // The actual value of the name, ie "='Sheet1'!A1"

    
};

} // namespace detail
} // namespace xlnt
