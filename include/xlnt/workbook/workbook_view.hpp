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

#include <cstddef>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

/// <summary>
/// A workbook can be opened in multiple windows with different views.
/// This class represents a particular view used by one window.
/// </summary>
class XLNT_CLASS workbook_view
{
public:
	bool auto_filter_date_grouping = false;
	bool minimized = false;
	bool show_horizontal_scroll = false;
	bool show_sheet_tabs = false;
	bool show_vertical_scroll = false;
	bool visible = true;

	std::size_t active_tab = 0;
	std::size_t first_sheet = 0;
	std::size_t tab_ratio = 500;
	std::size_t window_width = 28800;
	std::size_t window_height = 17460;
	std::size_t x_window = 0;
	std::size_t y_window = 460;
};

} // namespace xlnt
