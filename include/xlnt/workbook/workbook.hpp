// Copyright (c) 2014-2017 Thomas Fussell
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

#include <functional>
#include <iterator>
#include <memory>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

enum class calendar;
enum class relationship_type;

class alignment;
class border;
class calculation_properties;
class cell;
class cell_style;
class color;
class const_worksheet_iterator;
class drawing;
class fill;
class font;
class format;
class rich_text;
class manifest;
class named_range;
class number_format;
class path;
class pattern_fill;
class protection;
class range;
class range_reference;
class relationship;
class style;
class style_serializer;
class theme;
class workbook_view;
class worksheet;
class worksheet_iterator;
class zip_file;

struct datetime;

namespace detail {

struct stylesheet;
struct workbook_impl;
class xlsx_consumer;
class xlsx_producer;

} // namespace detail

/// <summary>
/// workbook is the container for all other parts of the document.
/// </summary>
class XLNT_API workbook
{
public:
    /// <summary>
    ///
    /// </summary>
    using iterator = worksheet_iterator;

    /// <summary>
    ///
    /// </summary>
    using const_iterator = const_worksheet_iterator;

    /// <summary>
    ///
    /// </summary>
    using reverse_iterator = std::reverse_iterator<iterator>;

    /// <summary>
    ///
    /// </summary>
    using const_reverse_iterator = std::reverse_iterator<const_iterator>;

    /// <summary>
    /// Swap the data held in workbooks "left" and "right".
    /// </summary>
    friend void swap(workbook &left, workbook &right);

    /// <summary>
    ///
    /// </summary>
    static workbook empty();

    // constructors

    /// <summary>
    /// Create a workbook containing a single empty worksheet.
    /// </summary>
    workbook();

    /// <summary>
    /// Move construct this workbook from existing workbook "other".
    /// </summary>
    workbook(workbook &&other);

    /// <summary>
    /// Copy construct this workbook from existing workbook "other".
    /// </summary>
    workbook(const workbook &other);

    /// <summary>
    /// Destroy this workbook.
    /// </summary>
    ~workbook();

    // add worksheets

    /// <summary>
    /// Create a sheet after the last sheet in this workbook and return it.
    /// </summary>
    worksheet create_sheet();

    /// <summary>
    /// Create a sheet at the specified index and return it.
    /// </summary>
    worksheet create_sheet(std::size_t index);

    /// <summary>
    /// This should be private...
    /// </summary>
    worksheet create_sheet_with_rel(const std::string &title, const relationship &rel);

    /// <summary>
    /// Create a new sheet initializing it with all of the data from the provided worksheet.
    /// </summary>
    void copy_sheet(worksheet worksheet);

    /// <summary>
    /// Create a new sheet at the specified index initializing it with all of the data
    /// from the provided worksheet.
    void copy_sheet(worksheet worksheet, std::size_t index);

    // get worksheets

    /// <summary>
    /// Returns the worksheet that was most recently accessed.
    /// This is also the sheet that will be shown when the workbook is opened
    /// in the spreadsheet editor program.
    /// </summary>
    worksheet active_sheet();

    /// <summary>
    /// Return the worksheet with the given name.
    /// This may throw an exception if the sheet isn't found.
    /// Use workbook::contains(const std::string &) to make sure the sheet exists.
    /// </summary>
    worksheet sheet_by_title(const std::string &sheet_name);

    /// <summary>
    /// Return the const worksheet with the given name.
    /// This may throw an exception if the sheet isn't found.
    /// Use workbook::contains(const std::string &) to make sure the sheet exists.
    /// </summary>
    const worksheet sheet_by_title(const std::string &sheet_name) const;

    /// <summary>
    /// Return the worksheet at the given index.
    /// </summary>
    worksheet sheet_by_index(std::size_t index);

    /// <summary>
    /// Return the const worksheet at the given index.
    /// </summary>
    const worksheet sheet_by_index(std::size_t index) const;

    /// <summary>
    /// Return the worksheet with a sheetId of id.
    /// </summary>
    worksheet sheet_by_id(std::size_t id);

    /// <summary>
    /// Return the const worksheet with a sheetId of id.
    /// </summary>
    const worksheet sheet_by_id(std::size_t id) const;

    /// <summary>
    /// Return true if this workbook contains a sheet with the given name.
    /// </summary>
    bool contains(const std::string &key) const;

    /// <summary>
    /// Return the index of the given worksheet.
    /// The worksheet must be owned by this workbook.
    /// </summary>
    std::size_t index(worksheet worksheet);

    // remove worksheets

    /// <summary>
    /// Remove the given worksheet from this workbook.
    /// </summary>
    void remove_sheet(worksheet worksheet);

    /// <summary>
    /// Delete every cell in this worksheet. After this is called, the
    /// worksheet will be equivalent to a newly created sheet at the same
    /// index and with the same title.
    /// </summary>
    void clear();

    // iterators

    /// <summary>
    /// Returns an iterator to the first worksheet in this workbook.
    /// </summary>
    iterator begin();

    /// <summary>
    /// Returns an iterator to the worksheet following the last worksheet of the workbook.
    /// This worksheet acts as a placeholder; attempting to access it will cause an
    /// exception to be thrown.
    /// </summary>
    iterator end();

    /// <summary>
    /// Returns a const iterator to the first worksheet in this workbook.
    /// </summary>
    const_iterator begin() const;

    /// <summary>
    /// Returns a const iterator to the worksheet following the last worksheet of the workbook.
    /// This worksheet acts as a placeholder; attempting to access it will cause an
    /// exception to be thrown.
    /// </summary>
    const_iterator end() const;

    /// <summary>
    /// Returns an iterator to the first worksheet in this workbook.
    /// </summary>
    const_iterator cbegin() const;

    /// <summary>
    /// Returns a const iterator to the worksheet following the last worksheet of the workbook.
    /// This worksheet acts as a placeholder; attempting to access it will cause an
    /// exception to be thrown.
    /// </summary>
    const_iterator cend() const;

    /// <summary>
    /// Apply the function "f" to every non-empty cell in every worksheet in this workbook.
    /// </summary>
    void apply_to_cells(std::function<void(cell)> f);

    /// <summary>
    /// Returns a temporary vector containing the titles of each sheet in the order
    /// of the sheets in the workbook.
    /// </summary>
    std::vector<std::string> sheet_titles() const;

    /// <summary>
    /// Returns the number of sheets in this workbook.
    /// </summary>
    std::size_t sheet_count() const;

    /// <summary>
    /// Returns true if the workbook has the core property with the given name.
    /// </summary>
    bool has_core_property(const std::string &property_name) const;

    /// <summary>
    ///
    /// </summary>
    template <typename T = std::string>
    T core_property(const std::string &property_name) const;

    /// <summary>
    ///
    /// </summary>
    template <typename T = std::string>
    void core_property(const std::string &property_name, const T value);

    /// <summary>
    /// Returns true if the workbook has the extended property with the given name.
    /// </summary>
    bool has_extended_property(const std::string &property_name) const;

    /// <summary>
    ///
    /// </summary>
    template <typename T = std::string>
    T extended_property(const std::string &property_name) const;

    /// <summary>
    ///
    /// </summary>
    template <typename T = std::string>
    void extended_property(const std::string &property_name, const T value);

    /// <summary>
    /// Returns true if the workbook has the custom property with the given name.
    /// </summary>
    bool has_custom_property(const std::string &property_name) const;

    /// <summary>
    ///
    /// </summary>
    template <typename T = std::string>
    T custom_property(const std::string &property_name) const;

    /// <summary>
    ///
    /// </summary>
    template <typename T = std::string>
    void custom_property(const std::string &property_name, const T value);

    /// <summary>
    ///
    /// </summary>
    calendar base_date() const;

    /// <summary>
    ///
    /// </summary>
    void base_date(calendar base_date);

    /// <summary>
    ///
    /// </summary>
    bool has_title() const;

    /// <summary>
    ///
    /// </summary>
    std::string title() const;

    /// <summary>
    ///
    /// </summary>
    void title(const std::string &title);

    // named ranges

    /// <summary>
    ///
    /// </summary>
    std::vector<xlnt::named_range> named_ranges() const;

    /// <summary>
    ///
    /// </summary>
    void create_named_range(const std::string &name, worksheet worksheet, const range_reference &reference);

    /// <summary>
    ///
    /// </summary>
    void create_named_range(const std::string &name, worksheet worksheet, const std::string &reference_string);

    /// <summary>
    ///
    /// </summary>
    bool has_named_range(const std::string &name) const;

    /// <summary>
    ///
    /// </summary>
    class range named_range(const std::string &name);

    /// <summary>
    ///
    /// </summary>
    void remove_named_range(const std::string &name);

    // serialization

    /// <summary>
    ///
    /// </summary>
    void save(std::vector<std::uint8_t> &data) const;

    /// <summary>
    ///
    /// </summary>
    void save(const std::string &filename) const;

    /// <summary>
    ///
    /// </summary>
    void save(const xlnt::path &filename) const;

    /// <summary>
    ///
    /// </summary>
    void save(std::ostream &stream) const;

    /// <summary>
    ///
    /// </summary>
    void save(const std::string &filename, const std::string &password);

    /// <summary>
    ///
    /// </summary>
    void save(const xlnt::path &filename, const std::string &password);

    /// <summary>
    ///
    /// </summary>
    void save(std::ostream &stream, const std::string &password);

    /// <summary>
    ///
    /// </summary>
    void save(const std::vector<std::uint8_t> &data, const std::string &password);

    /// <summary>
    ///
    /// </summary>
    void load(const std::vector<std::uint8_t> &data);

    /// <summary>
    ///
    /// </summary>
    void load(const std::string &filename);

    /// <summary>
    ///
    /// </summary>
    void load(const xlnt::path &filename);

    /// <summary>
    ///
    /// </summary>
    void load(std::istream &stream);

    /// <summary>
    ///
    /// </summary>
    void load(const std::string &filename, const std::string &password);

    /// <summary>
    ///
    /// </summary>
    void load(const xlnt::path &filename, const std::string &password);

    /// <summary>
    ///
    /// </summary>
    void load(std::istream &stream, const std::string &password);

    /// <summary>
    ///
    /// </summary>
    void load(const std::vector<std::uint8_t> &data, const std::string &password);

#ifdef _MSC_VER
    /// <summary>
    ///
    /// </summary>
    void save(const std::wstring &filename);

    /// <summary>
    ///
    /// </summary>
    void save(const std::wstring &filename, const std::string &password);

    /// <summary>
    ///
    /// </summary>
    void load(const std::wstring &filename);

    /// <summary>
    ///
    /// </summary>
    void load(const std::wstring &filename, const std::string &password);
#endif

    /// <summary>
    ///
    /// </summary>
    bool has_view() const;

    /// <summary>
    ///
    /// </summary>
    workbook_view view() const;

    /// <summary>
    ///
    /// </summary>
    void view(const workbook_view &view);

    /// <summary>
    ///
    /// </summary>
    bool has_code_name() const;

    /// <summary>
    ///
    /// </summary>
    std::string code_name() const;

    /// <summary>
    ///
    /// </summary>
    void code_name(const std::string &code_name);

    /// <summary>
    ///
    /// </summary>
    bool has_file_version() const;

    /// <summary>
    ///
    /// </summary>
    std::string app_name() const;

    /// <summary>
    ///
    /// </summary>
    std::size_t last_edited() const;

    /// <summary>
    ///
    /// </summary>
    std::size_t lowest_edited() const;

    /// <summary>
    ///
    /// </summary>
    std::size_t rup_build() const;

    // theme

    /// <summary>
    ///
    /// </summary>
    bool has_theme() const;

    /// <summary>
    ///
    /// </summary>
    const xlnt::theme &theme() const;

    /// <summary>
    ///
    /// </summary>
    void theme(const class theme &value);

    // formats

    /// <summary>
    ///
    /// </summary>
    xlnt::format format(std::size_t format_index);

    /// <summary>
    ///
    /// </summary>
    const xlnt::format format(std::size_t format_index) const;

    /// <summary>
    ///
    /// </summary>
    xlnt::format create_format(bool default_format = false);

    /// <summary>
    ///
    /// </summary>
    void clear_formats();

    // styles

    /// <summary>
    ///
    /// </summary>
    bool has_style(const std::string &name) const;

    /// <summary>
    ///
    /// </summary>
    class style style(const std::string &name);

    /// <summary>
    ///
    /// </summary>
    const class style style(const std::string &name) const;

    /// <summary>
    ///
    /// </summary>
    class style create_style(const std::string &name);

    /// <summary>
    ///
    /// </summary>
    void clear_styles();

    // manifest

    /// <summary>
    ///
    /// </summary>
    class manifest &manifest();

    /// <summary>
    ///
    /// </summary>
    const class manifest &manifest() const;

    // shared strings

    /// <summary>
    ///
    /// </summary>
    void add_shared_string(const rich_text &shared, bool allow_duplicates = false);

    /// <summary>
    ///
    /// </summary>
    std::vector<rich_text> &shared_strings();

    /// <summary>
    ///
    /// </summary>
    const std::vector<rich_text> &shared_strings() const;

    // thumbnail

    /// <summary>
    ///
    /// </summary>
    void thumbnail(const std::vector<std::uint8_t> &thumbnail,
        const std::string &extension, const std::string &content_type);

    /// <summary>
    ///
    /// </summary>
    const std::vector<std::uint8_t> &thumbnail() const;

    // calculation properties

    /// <summary>
    ///
    /// </summary>
    bool has_calculation_properties() const;

    /// <summary>
    ///
    /// </summary>
    class calculation_properties calculation_properties() const;

    /// <summary>
    ///
    /// </summary>
    void calculation_properties(const class calculation_properties &props);

    // operators

    /// <summary>
    /// Set the contents of this workbook to be equal to those of "other".
    /// Other is passed as value to allow for copy-swap idiom.
    /// </summary>
    workbook &operator=(workbook other);

    /// <summary>
    /// Return the worksheet with a title of "name".
    /// </summary>
    worksheet operator[](const std::string &name);

    /// <summary>
    /// Return the worksheet at "index".
    /// </summary>
    worksheet operator[](std::size_t index);

    /// <summary>
    /// Return true if this workbook internal implementation points to the same
    /// memory as rhs's.
    /// </summary>
    bool operator==(const workbook &rhs) const;

    /// <summary>
    /// Return true if this workbook internal implementation doesn't point to the same
    /// memory as rhs's.
    /// </summary>
    bool operator!=(const workbook &rhs) const;

private:
    friend class worksheet;
    friend class detail::xlsx_consumer;
    friend class detail::xlsx_producer;

    /// <summary>
    ///
    /// </summary>
    workbook(detail::workbook_impl *impl);

    /// <summary>
    ///
    /// </summary>
    detail::workbook_impl &impl();

    /// <summary>
    ///
    /// </summary>
    const detail::workbook_impl &impl() const;

    /// <summary>
    ///
    /// </summary>
    void register_app_properties_in_manifest();

    /// <summary>
    ///
    /// </summary>
    void register_core_properties_in_manifest();

    /// <summary>
    ///
    /// </summary>
    void register_shared_string_table_in_manifest();

    /// <summary>
    ///
    /// </summary>
    void register_stylesheet_in_manifest();

    /// <summary>
    ///
    /// </summary>
    void register_theme_in_manifest();

    /// <summary>
    ///
    /// </summary>
    void register_comments_in_manifest(worksheet ws);

    /// <summary>
    /// Removes calcChain part from manifest if no formulae remain in workbook.
    /// </summary>
    void garbage_collect_formulae();

    /// <summary>
    /// An opaque pointer to a structure that holds all of the data relating to this workbook.
    /// </summary>
    std::unique_ptr<detail::workbook_impl> d_;
};

} // namespace xlnt
