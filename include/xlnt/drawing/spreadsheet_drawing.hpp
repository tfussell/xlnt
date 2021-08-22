// Copyright (c) 2018
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

#include <xlnt/xlnt_config.hpp>
#include <string>
#include <vector>

namespace xml {
class parser;
class serializer;
} // namespace xml

namespace xlnt {

class worksheet;

namespace drawing {

/// <summary>
/// The spreadsheet_drawing class encapsulates the information
/// captured from objects within the spreadsheetDrawing schema.
/// </summary>
class XLNT_API spreadsheet_drawing
{
public:
    spreadsheet_drawing(xml::parser &parser);
    void serialize(xml::serializer &serializer);

    std::vector<std::string> get_embed_ids();

private:
    std::string serialized_value_;
    std::vector<std::string> embed_ids_;
};

} // namespace drawing
} // namespace xlnt
