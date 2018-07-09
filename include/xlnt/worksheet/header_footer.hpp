// Copyright (c) 2014-2018 Thomas Fussell
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

#include <cstdint>
#include <string>
#include <unordered_map>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/rich_text.hpp>
#include <xlnt/utils/scoped_enum_hash.hpp>

namespace xlnt {

/// <summary>
/// Represents the header and footer of a sheet in a workbook.
/// </summary>
class XLNT_API header_footer
{
public:
    /// <summary>
    /// Enumerates the three possible locations of a header or footer.
    /// </summary>
    enum class location
    {
        left,
        center,
        right
    };

    // General Properties

    /// <summary>
    /// True if any text has been added for a header at any location on any page.
    /// </summary>
    bool has_header() const;

    /// <summary>
    /// True if any text has been added for a footer at any location on any page.
    /// </summary>
    bool has_footer() const;

    /// <summary>
    /// True if headers and footers should align to the page margins.
    /// </summary>
    bool align_with_margins() const;

    /// <summary>
    /// Set to true if headers and footers should align to the page margins.
    /// Set to false if headers and footers should align to the edge of the page.
    /// </summary>
    header_footer &align_with_margins(bool align);

    /// <summary>
    /// True if headers and footers differ based on page number.
    /// </summary>
    bool different_odd_even() const;

    /// <summary>
    /// True if headers and footers are different on the first page.
    /// </summary>
    bool different_first() const;

    /// <summary>
    /// True if headers and footers should scale to match the worksheet.
    /// </summary>
    bool scale_with_doc() const;

    /// <summary>
    /// Set to true if headers and footers should scale to match the worksheet.
    /// </summary>
    header_footer &scale_with_doc(bool scale);


    // Normal Header

    /// <summary>
    /// True if any text has been added at the given location on any page.
    /// </summary>
    bool has_header(location where) const;

    /// <summary>
    /// Remove all headers from all pages.
    /// </summary>
    void clear_header();

    /// <summary>
    /// Remove header at the given location on any page.
    /// </summary>
    void clear_header(location where);

    /// <summary>
    /// Add a header at the given location with the given text.
    /// </summary>
    header_footer &header(location where, const std::string &text);

    /// <summary>
    /// Add a header at the given location with the given text.
    /// </summary>
    header_footer &header(location where, const rich_text &text);

    /// <summary>
    /// Get the text of the  header at the given location. If headers are
    /// different on odd and even pages, the odd header will be returned.
    /// </summary>
    rich_text header(location where) const;

    // First Page Header

    /// <summary>
    /// True if a header has been set for the first page at any location.
    /// </summary>
    bool has_first_page_header() const;

    /// <summary>
    /// True if a header has been set for the first page at the given location.
    /// </summary>
    bool has_first_page_header(location where) const;

    /// <summary>
    /// Remove all headers from the first page.
    /// </summary>
    void clear_first_page_header();

    /// <summary>
    /// Remove header from the first page at the given location.
    /// </summary>
    void clear_first_page_header(location where);

    /// <summary>
    /// Add a header on the first page at the given location with the given text.
    /// </summary>
    header_footer &first_page_header(location where, const rich_text &text);

    /// <summary>
    /// Get the text of the first page header at the given location. If no first
    /// page header has been set, the general header for that location will
    /// be returned.
    /// </summary>
    rich_text first_page_header(location where) const;

    // Odd/Even Header

    /// <summary>
    /// True if different headers have been set for odd and even pages.
    /// </summary>
    bool has_odd_even_header() const;

    /// <summary>
    /// True if different headers have been set for odd and even pages at the given location.
    /// </summary>
    bool has_odd_even_header(location where) const;

    /// <summary>
    /// Remove odd/even headers at all locations.
    /// </summary>
    void clear_odd_even_header();

    /// <summary>
    /// Remove odd/even headers at the given location.
    /// </summary>
    void clear_odd_even_header(location where);

    /// <summary>
    /// Add a header for odd pages at the given location with the given text.
    /// </summary>
    header_footer &odd_even_header(location where, const rich_text &odd, const rich_text &even);

    /// <summary>
    /// Get the text of the odd page header at the given location. If no odd
    /// page header has been set, the general header for that location will
    /// be returned.
    /// </summary>
    rich_text odd_header(location where) const;

    /// <summary>
    /// Get the text of the even page header at the given location. If no even
    /// page header has been set, the general header for that location will
    /// be returned.
    /// </summary>
    rich_text even_header(location where) const;


    // Normal Footer

    /// <summary>
    /// True if any text has been added at the given location on any page.
    /// </summary>
    bool has_footer(location where) const;

    /// <summary>
    /// Remove all footers from all pages.
    /// </summary>
    void clear_footer();

    /// <summary>
    /// Remove footer at the given location on any page.
    /// </summary>
    void clear_footer(location where);

    /// <summary>
    /// Add a footer at the given location with the given text.
    /// </summary>
    header_footer &footer(location where, const std::string &text);

    /// <summary>
    /// Add a footer at the given location with the given text.
    /// </summary>
    header_footer &footer(location where, const rich_text &text);

    /// <summary>
    /// Get the text of the  footer at the given location. If footers are
    /// different on odd and even pages, the odd footer will be returned.
    /// </summary>
    rich_text footer(location where) const;

    // First Page footer

    /// <summary>
    /// True if a footer has been set for the first page at any location.
    /// </summary>
    bool has_first_page_footer() const;

    /// <summary>
    /// True if a footer has been set for the first page at the given location.
    /// </summary>
    bool has_first_page_footer(location where) const;

    /// <summary>
    /// Remove all footers from the first page.
    /// </summary>
    void clear_first_page_footer();

    /// <summary>
    /// Remove footer from the first page at the given location.
    /// </summary>
    void clear_first_page_footer(location where);

    /// <summary>
    /// Add a footer on the first page at the given location with the given text.
    /// </summary>
    header_footer &first_page_footer(location where, const rich_text &text);

    /// <summary>
    /// Get the text of the first page footer at the given location. If no first
    /// page footer has been set, the general footer for that location will
    /// be returned.
    /// </summary>
    rich_text first_page_footer(location where) const;


    // Odd/Even Footer

    /// <summary>
    /// True if different footers have been set for odd and even pages.
    /// </summary>
    bool has_odd_even_footer() const;

    /// <summary>
    /// True if different footers have been set for odd and even pages at the given location.
    /// </summary>
    bool has_odd_even_footer(location where) const;

    /// <summary>
    /// Remove odd/even footers at all locations.
    /// </summary>
    void clear_odd_even_footer();

    /// <summary>
    /// Remove odd/even footers at the given location.
    /// </summary>
    void clear_odd_even_footer(location where);

    /// <summary>
    /// Add a footer for odd pages at the given location with the given text.
    /// </summary>
    header_footer &odd_even_footer(location where, const rich_text &odd, const rich_text &even);

    /// <summary>
    /// Get the text of the odd page footer at the given location. If no odd
    /// page footer has been set, the general footer for that location will
    /// be returned.
    /// </summary>
    rich_text odd_footer(location where) const;

    /// <summary>
    /// Get the text of the even page footer at the given location. If no even
    /// page footer has been set, the general footer for that location will
    /// be returned.
    /// </summary>
    rich_text even_footer(location where) const;

    bool operator==(const header_footer &rhs) const;

private:
    bool align_with_margins_ = false;
    bool different_odd_even_ = false;
    bool scale_with_doc_ = false;

    using container = std::unordered_map<location, rich_text, scoped_enum_hash<location>>;

    container odd_headers_;
    container even_headers_;
    container first_headers_;
    container odd_footers_;
    container even_footers_;
    container first_footers_;
};

} // namespace xlnt
