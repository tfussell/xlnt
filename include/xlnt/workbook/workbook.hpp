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

#include <functional>
#include <iterator>
#include <memory>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/packaging/relationship.hpp>

namespace xlnt {

class alignment;
class app_properties;
class border;
class cell;
class cell_style;
class color;
class const_worksheet_iterator;
class document_properties;
class drawing;
class fill;
class font;
class format;
class manifest;
class named_range;
class number_format;
class pattern_fill;
class protection;
class range;
class range_reference;
class relationship;
class style;
class style_serializer;
class text;
class theme;
class worksheet;
class worksheet_iterator;
class zip_file;

namespace detail {
struct workbook_impl;
} // namespace detail

/// <summary>
/// workbook is the container for all other parts of the document.
/// </summary>
class XLNT_CLASS workbook
{
public:
    using iterator = worksheet_iterator;
    using const_iterator = const_worksheet_iterator;
    using reverse_iterator = std::reverse_iterator<iterator>;
    using const_reverse_iterator = std::reverse_iterator<const_iterator>;

    /// <summary>
    /// Swap the data held in workbooks "left" and "right".
    /// </summary>
    friend void swap(workbook &left, workbook &right);

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

    // general properties

    bool get_guess_types() const;
    void set_guess_types(bool guess);

    bool get_data_only() const;
    void set_data_only(bool data_only);
    
    bool get_read_only() const;
    void set_read_only(bool read_only);

    // add worksheets

    worksheet create_sheet();
    worksheet create_sheet(std::size_t index);
    worksheet create_sheet(const std::string &title);
    worksheet create_sheet(std::size_t index, const std::string &title);
    worksheet create_sheet(const std::string &title, const relationship &rel);

    void copy_sheet(worksheet worksheet);
    void copy_sheet(worksheet worksheet, std::size_t index);

    // get worksheets

    /// <summary>
    /// Returns the worksheet that was most recently accessed.
    /// This is also the sheet that will be shown when the workbook is opened
    /// in the spreadsheet editor program.
    /// </summary>
    worksheet get_active_sheet();

    /// <summary>
    /// Return the worksheet with the given name.
    /// This may throw an exception if the sheet isn't found.
    /// Use workbook::contains(const std::string &) to make sure the sheet exists.
    /// </summary>
    worksheet get_sheet_by_name(const std::string &sheet_name);
    
    /// <summary>
    /// Return the const worksheet with the given name.
    /// This may throw an exception if the sheet isn't found.
    /// Use workbook::contains(const std::string &) to make sure the sheet exists.
    /// </summary>
    const worksheet get_sheet_by_name(const std::string &sheet_name) const;

    /// <summary>
    /// Return the worksheet at the given index.
    /// </summary>
    worksheet get_sheet_by_index(std::size_t index);

    /// <summary>
    /// Return the const worksheet at the given index.
    /// </summary>
    const worksheet get_sheet_by_index(std::size_t index) const;

    /// <summary>
    /// Return true if this workbook contains a sheet with the given name.
    /// </summary>
    bool contains(const std::string &key) const;

    /// <summary>
    /// Return the index of the given worksheet.
    /// The worksheet must be owned by this workbook.
    /// </summary>
    std::size_t get_index(worksheet worksheet);

    // remove worksheets

    void remove_sheet(worksheet worksheet);
    void clear();

    // iterators

    iterator begin();
    iterator end();

    const_iterator begin() const;
    const_iterator end() const;

    const_iterator cbegin() const;
    const_iterator cend() const;
    
    reverse_iterator rbegin();
    reverse_iterator rend();

    const_reverse_iterator rbegin() const;
    const_reverse_iterator rend() const;

    const_reverse_iterator crbegin() const;
    const_reverse_iterator crend() const;

    std::vector<std::string> get_sheet_names() const;

    document_properties &get_properties();
    const document_properties &get_properties() const;

    app_properties &get_app_properties();
    const app_properties &get_app_properties() const;

    // named ranges

    std::vector<named_range> get_named_ranges() const;
    void create_named_range(const std::string &name, worksheet worksheet, const range_reference &reference);
    void create_named_range(const std::string &name, worksheet worksheet, const std::string &reference_string);
    bool has_named_range(const std::string &name) const;
    range get_named_range(const std::string &name);
    void remove_named_range(const std::string &name);

    // serialization

    bool save(std::vector<unsigned char> &data);
    bool save(const std::string &filename);
    bool load(const std::vector<unsigned char> &data);
    bool load(const std::string &filename);
    bool load(std::istream &stream);
    bool load(zip_file &archive);

    void set_code_name(const std::string &code_name);

    // theme

    bool has_loaded_theme() const;
    const theme &get_loaded_theme() const;

    // formats
    
    format &get_format(std::size_t format_index);
    const format &get_format(std::size_t format_index) const;
    std::size_t add_format(const format &new_format);
    void clear_formats();
    std::vector<format> &get_formats();
    const std::vector<format> &get_formats() const;

    // named styles

    bool has_style(const std::string &name) const;
    style &get_style(const std::string &name);
    const style &get_style(const std::string &name) const;
    style &get_style_by_id(std::size_t style_id);
    const style &get_style_by_id(std::size_t style_id) const;
    std::size_t get_style_id(const std::string &name) const;
    style &create_style(const std::string &name);
    std::size_t add_style(const style &new_style);
    void clear_styles();
    const std::vector<style> &get_styles() const;

    // manifest

    manifest &get_manifest();
    const manifest &get_manifest() const;

    // relationships

    void create_relationship(const std::string &id, const std::string &target, relationship::type type);
    relationship get_relationship(const std::string &id) const;
    const std::vector<relationship> &get_relationships() const;

    void create_root_relationship(const std::string &id, const std::string &target, relationship::type type);
    const std::vector<relationship> &get_root_relationships() const;

    // shared strings

    void add_shared_string(const text &shared, bool allow_duplicates=false);
    std::vector<text> &get_shared_strings();
    const std::vector<text> &get_shared_strings() const;
    
    // thumbnail

    void set_thumbnail(const std::vector<std::uint8_t> &thumbnail);
    const std::vector<std::uint8_t> &get_thumbnail() const;

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
    friend class excel_serializer;
    friend class worksheet;

    /// <summary>
    /// Helper function to calculate an index from a worksheet filename.
    /// </summary>
    static std::size_t index_from_ws_filename(const std::string &filename);

    /// <summary>
    /// Get the name of the next unused relationship.
    /// </summary>
    std::string next_relationship_id() const;

    /// <summary>
    /// Apply the function "f" to every cell in every worksheet in this workbook.
    /// </summary>
    void apply_to_cells(std::function<void(cell)> f);
    
    /// <summary>
    /// An opaque pointer to a structure that holds all of the data relating to this workbook.
    /// </summary>
    std::unique_ptr<detail::workbook_impl> d_;
};

} // namespace xlnt
