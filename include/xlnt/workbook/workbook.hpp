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

namespace xlnt {

class alignment;
class border;
class cell;
class cell_style;
class color;
class const_worksheet_iterator;
class drawing;
class fill;
class font;
class format;
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
class text;
class theme;
class workbook_view;
class worksheet;
class worksheet_iterator;
class zip_file;

struct datetime;

enum class calendar;
enum class relationship_type;

namespace detail {
struct stylesheet;
struct workbook_impl;
class xlsx_consumer;
class xlsx_producer;
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

	static workbook minimal();

	static workbook empty_excel();

	static workbook empty_libre_office();

	static workbook empty_numbers();

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

	/// <summary>
	/// Returns true if guess_types is enabled for this workbook.
	/// </summary>
    bool get_guess_types() const;

	/// <summary>
	/// Set to true to guess the type represented by a string set as the value
	/// for a cell and then set the actual value of the cell to that type.
	/// For example, cell.set_value("1") with guess_types enabled will set
	/// type of the cell to numeric and it's value to the number 1.
	/// </summary>
    void set_guess_types(bool guess);

	/// <summary>
	/// ?
	/// </summary>
    bool get_data_only() const;

	/// <summary>
	/// ?
	/// </summary>
    void set_data_only(bool data_only);

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
    worksheet get_active_sheet();

    /// <summary>
    /// Return the worksheet with the given name.
    /// This may throw an exception if the sheet isn't found.
    /// Use workbook::contains(const std::string &) to make sure the sheet exists.
    /// </summary>
    worksheet get_sheet_by_title(const std::string &sheet_name);
    
    /// <summary>
    /// Return the const worksheet with the given name.
    /// This may throw an exception if the sheet isn't found.
    /// Use workbook::contains(const std::string &) to make sure the sheet exists.
    /// </summary>
    const worksheet get_sheet_by_title(const std::string &sheet_name) const;

    /// <summary>
    /// Return the worksheet at the given index.
    /// </summary>
    worksheet get_sheet_by_index(std::size_t index);

    /// <summary>
    /// Return the const worksheet at the given index.
    /// </summary>
    const worksheet get_sheet_by_index(std::size_t index) const;

	/// <summary>
	/// Return the worksheet with a sheetId of id.
	/// </summary>
	worksheet get_sheet_by_id(std::size_t id);

	/// <summary>
	/// Return the const worksheet with a sheetId of id.
	/// </summary>
	const worksheet get_sheet_by_id(std::size_t id) const;

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
	/// Returns a temporary vector containing the titles of each sheet in the order
	/// of the sheets in the workbook.
	/// </summary>
    std::vector<std::string> get_sheet_titles() const;

	/// <summary>
	/// Returns the number of sheets in this workbook.
	/// </summary>
	std::size_t get_sheet_count() const;

	/// <summary>
	/// Returns the name of the application that created this workbook.
	/// </summary>
	std::string get_application() const;

	/// <summary>
	/// Sets the name of the application that created this workbook to "application".
	/// </summary>
	void set_application(const std::string &application);

	calendar get_base_date() const;
	void set_base_date(calendar base_date);

	std::string get_creator() const;
	void set_creator(const std::string &creator);

	std::string get_last_modified_by() const;
	void set_last_modified_by(const std::string &last_modified_by);

	datetime get_created() const;
	void set_created(const datetime &when);

	datetime get_modified() const;
	void set_modified(const datetime &when);

	int get_doc_security() const;
	void set_doc_security(int doc_security);

	bool get_scale_crop() const;
	void set_scale_crop(bool scale_crop);

	std::string get_company() const;
	void set_company(const std::string &company);

	bool links_up_to_date() const;
	void set_links_up_to_date(bool links_up_to_date);

	bool is_shared_doc() const;
	void set_shared_doc(bool shared_doc);

	bool hyperlinks_changed() const;
	void set_hyperlinks_changed(bool hyperlinks_changed);

	std::string get_app_version() const;
	void set_app_version(const std::string &version);

	std::string get_title() const;
	void set_title(const std::string &title);

    // named ranges

    std::vector<named_range> get_named_ranges() const;
    void create_named_range(const std::string &name, worksheet worksheet, const range_reference &reference);
    void create_named_range(const std::string &name, worksheet worksheet, const std::string &reference_string);
    bool has_named_range(const std::string &name) const;
    range get_named_range(const std::string &name);
    void remove_named_range(const std::string &name);

    // serialization

	void save(std::vector<std::uint8_t> &data) const;
	void save(const std::string &filename) const;
	void save(const xlnt::path &filename) const;
	void save(std::ostream &stream) const;

	void load(const std::vector<std::uint8_t> &data);
	void load(const std::string &filename);
	void load(const xlnt::path &filename);
	void load(std::istream &stream);

	bool has_view() const;
	workbook_view get_view() const;
	void set_view(const workbook_view &view);

	bool has_code_name() const;
	std::string get_code_name() const;
    void set_code_name(const std::string &code_name);

	bool x15_enabled() const;
	void enable_x15();
	void disable_x15();

	bool has_absolute_path() const;
	path get_absolute_path() const;
	void set_absolute_path(const path &absolute_path);
	void clear_absolute_path();

	bool has_properties() const;

	bool has_file_version() const;
	std::string get_app_name() const;
	std::size_t get_last_edited() const;
	std::size_t get_lowest_edited() const;
	std::size_t get_rup_build() const;

	bool has_arch_id() const;

	// calculation

	bool has_calculation_properties() const;

    // theme

    bool has_theme() const;
    const theme &get_theme() const;
	void set_theme(const theme &value);

    // formats
    
    format &get_format(std::size_t format_index);
    const format &get_format(std::size_t format_index) const;
	format &create_format();
    void clear_formats();

    // styles

    bool has_style(const std::string &name) const;
    style &get_style(const std::string &name);
    const style &get_style(const std::string &name) const;
    style &create_style(const std::string &name);
    void clear_styles();

    // manifest

    manifest &get_manifest();
    const manifest &get_manifest() const;

    // shared strings

    void add_shared_string(const text &shared, bool allow_duplicates=false);
    std::vector<text> &get_shared_strings();
    const std::vector<text> &get_shared_strings() const;
    
    // thumbnail

    void set_thumbnail(const std::vector<std::uint8_t> &thumbnail, 
		const std::string &extension, const std::string &content_type);
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
	friend class worksheet;
	friend class detail::xlsx_consumer;
	friend class detail::xlsx_producer;

	workbook(detail::workbook_impl *impl);

	detail::workbook_impl &impl();

	const detail::workbook_impl &impl() const;

    /// <summary>
    /// Apply the function "f" to every cell in every worksheet in this workbook.
    /// </summary>
    void apply_to_cells(std::function<void(cell)> f);

	void register_app_properties_in_manifest();

	void register_core_properties_in_manifest();

	void register_shared_string_table_in_manifest();

	void register_stylesheet_in_manifest();

	void register_theme_in_manifest();
    
    /// <summary>
    /// An opaque pointer to a structure that holds all of the data relating to this workbook.
    /// </summary>
    std::unique_ptr<detail::workbook_impl> d_;
};

} // namespace xlnt
