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

#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/hash_combine.hpp>
#include <xlnt/styles/horizontal_alignment.hpp>
#include <xlnt/styles/vertical_alignment.hpp>
#include <xlnt/utils/hashable.hpp>

namespace xlnt {

/// <summary>
/// Alignment options for use in styles.
/// </summary>
class XLNT_CLASS alignment : public hashable
{
public:
    bool get_wrap_text() const;

    void set_wrap_text(bool wrap_text);

    bool has_horizontal() const;

    horizontal_alignment get_horizontal() const;

    void set_horizontal(horizontal_alignment horizontal);

    bool has_vertical() const;

    vertical_alignment get_vertical() const;

    void set_vertical(vertical_alignment vertical);
    
protected:
    std::string to_hash_string() const override;

private:
    horizontal_alignment horizontal_ = horizontal_alignment::none;
    vertical_alignment vertical_ = vertical_alignment::none;
    int text_rotation_ = 0;
    bool wrap_text_ = false;
    bool shrink_to_fit_ = false;
    int indent_ = 0;
};

} // namespace xlnt
