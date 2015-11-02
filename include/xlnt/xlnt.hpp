// Copyright (c) 2015 Thomas Fussell
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

const std::string xlnt_version = "0.1.0";

const std::string author = "Thomas Fussell";
const std::string license = "MIT";
const std::string author_email = "thomas.fussell@gmail.com";
const std::string url = "https://github.com/tfussell/xlnt";
const std::string download_url = "https://github.com/tfussell/xlnt/archive/master.zip";

#include <xlnt/config.hpp>

#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/cell/comment.hpp>
#include <xlnt/common/datetime.hpp>
#include <xlnt/common/encoding.hpp>
#include <xlnt/common/exceptions.hpp>
#include <xlnt/common/relationship.hpp>
#include <xlnt/common/types.hpp>
#include <xlnt/common/zip_file.hpp>
#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/color.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/named_style.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/styles/side.hpp>
#include <xlnt/styles/style.hpp>
#include <xlnt/workbook/document_properties.hpp>
#include <xlnt/workbook/document_security.hpp>
#include <xlnt/workbook/external_book.hpp>
#include <xlnt/workbook/manifest.hpp>
#include <xlnt/workbook/named_range.hpp>
#include <xlnt/workbook/theme.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/cell_vector.hpp>
#include <xlnt/worksheet/column_properties.hpp>
#include <xlnt/worksheet/major_order.hpp>
#include <xlnt/worksheet/page_margins.hpp>
#include <xlnt/worksheet/page_setup.hpp>
#include <xlnt/worksheet/pane.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/row_properties.hpp>
#include <xlnt/worksheet/selection.hpp>
#include <xlnt/worksheet/sheet_protection.hpp>
#include <xlnt/worksheet/sheet_view.hpp>
#include <xlnt/worksheet/worksheet_properties.hpp>
#include <xlnt/worksheet/worksheet.hpp>
