// Copyright (c) 2017-2018 Thomas Fussell
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

namespace xlnt {

/// <summary>
/// Every core property in a workbook must be one of these types.
/// </summary>
enum class core_property
{
    category,
    content_status,
    created,
    creator,
    description,
    identifier,
    keywords,
    language,
    last_modified_by,
    last_printed,
    modified,
    revision,
    subject,
    title,
    version
};

/// <summary>
/// Every extended property in a workbook must be one of these types.
/// </summary>
enum class extended_property
{
    application,
    app_version,
    characters,
    characters_with_spaces,
    company,
    dig_sig,
    doc_security,
    heading_pairs,
    hidden_slides,
    h_links,
    hyperlink_base,
    hyperlinks_changed,
    lines,
    links_up_to_date,
    manager,
    m_m_clips,
    notes,
    pages,
    paragraphs,
    presentation_format,
    scale_crop,
    shared_doc,
    slides,
    template_,
    titles_of_parts,
    total_time,
    words
};

} // namespace xlnt
