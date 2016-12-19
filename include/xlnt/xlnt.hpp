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

#include <xlnt/xlnt_config.hpp>

// cell
#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/cell/cell_type.hpp>
#include <xlnt/cell/comment.hpp>
#include <xlnt/cell/formatted_text.hpp>
#include <xlnt/cell/index_types.hpp>
#include <xlnt/cell/text_run.hpp>

// packaging
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/packaging/relationship.hpp>

// styles
#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/border_style.hpp>
#include <xlnt/styles/color.hpp>
#include <xlnt/styles/diagonal_direction.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/format.hpp>
#include <xlnt/styles/horizontal_alignment.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/styles/style.hpp>
#include <xlnt/styles/vertical_alignment.hpp>

// utils
#include <xlnt/utils/calendar.hpp>
#include <xlnt/utils/date.hpp>
#include <xlnt/utils/datetime.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/utils/path.hpp>
#include <xlnt/utils/time.hpp>
#include <xlnt/utils/timedelta.hpp>

// workbook
#include <xlnt/workbook/const_worksheet_iterator.hpp>
#include <xlnt/workbook/document_security.hpp>
#include <xlnt/workbook/external_book.hpp>
#include <xlnt/workbook/named_range.hpp>
#include <xlnt/workbook/theme.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/workbook/worksheet_iterator.hpp>

// worksheet
#include <xlnt/worksheet/cell_iterator.hpp>
#include <xlnt/worksheet/cell_vector.hpp>
#include <xlnt/worksheet/column_properties.hpp>
#include <xlnt/worksheet/const_cell_iterator.hpp>
#include <xlnt/worksheet/const_range_iterator.hpp>
#include <xlnt/worksheet/header_footer.hpp>
#include <xlnt/worksheet/header_footer.hpp>
#include <xlnt/worksheet/major_order.hpp>
#include <xlnt/worksheet/orientation.hpp>
#include <xlnt/worksheet/page_break.hpp>
#include <xlnt/worksheet/page_margins.hpp>
#include <xlnt/worksheet/page_setup.hpp>
#include <xlnt/worksheet/pane.hpp>
#include <xlnt/worksheet/paper_size.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/range_iterator.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/row_properties.hpp>
#include <xlnt/worksheet/selection.hpp>
#include <xlnt/worksheet/sheet_protection.hpp>
#include <xlnt/worksheet/sheet_state.hpp>
#include <xlnt/worksheet/sheet_view.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/worksheet/worksheet_properties.hpp>
