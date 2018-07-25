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

#include <functional>
#include <iterator>
#include <map>
#include <memory>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/cell/rich_text.hpp>

namespace xlnt {

enum class calendar;
enum class core_property;
enum class extended_property;
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
class metadata_property;
class named_range;
class number_format;
class path;
class pattern_fill;
class protection;
class range;
class range_reference;
class relationship;
class streaming_workbook_reader;
class style;
class style_serializer;
class theme;
class variant;
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
    /// typedef for the iterator used for iterating through this workbook
    /// (non-const) in a range-based for loop.
    /// </summary>
    using iterator = worksheet_iterator;

    /// <summary>
    /// typedef for the iterator used for iterating through this workbook
    /// (const) in a range-based for loop.
    /// </summary>
    using const_iterator = const_worksheet_iterator;

    /// <summary>
    /// typedef for the iterator used for iterating through this workbook
    /// (non-const) in a range-based for loop in reverse order using
    /// std::make_reverse_iterator.
    /// </summary>
    using reverse_iterator = std::reverse_iterator<iterator>;

    /// <summary>
    /// typedef for the iterator used for iterating through this workbook
    /// (const) in a range-based for loop in reverse order using
    /// std::make_reverse_iterator.
    /// </summary>
    using const_reverse_iterator = std::reverse_iterator<const_iterator>;

    /// <summary>
    /// Constructs and returns an empty workbook similar to a default.
    /// Excel workbook
    /// </summary>
    static workbook empty();

    // Constructors

    /// <summary>
    /// Default constructor. Constructs a workbook containing a single empty
    /// worksheet using workbook::empty().
    /// </summary>
    workbook();

    /// <summary>
    /// load the xlsx file at path
    /// </summary>
    workbook(const xlnt::path &file);

    /// <summary>
    /// load the encrpyted xlsx file at path
    /// </summary>
    workbook(const xlnt::path &file, const std::string& password);

    /// <summary>
    /// construct the workbook from any data stream where the data is the binary form of a workbook
    /// </summary>
    workbook(std::istream & data);

    /// <summary>
    /// construct the workbook from any data stream where the data is the binary form of an encrypted workbook
    /// </summary>
    workbook(std::istream &data, const std::string& password);

    /// <summary>
    /// Move constructor. Constructs a workbook from existing workbook, other.
    /// </summary>
    workbook(workbook &&other);

    /// <summary>
    /// Copy constructor. Constructs this workbook from existing workbook, other.
    /// </summary>
    workbook(const workbook &other);

    /// <summary>
    /// Destroys this workbook, deallocating all internal storage space. Any pimpl
    /// wrapper classes (e.g. cell) pointing into this workbook will be invalid
    /// after this is executed.
    /// </summary>
    ~workbook();

    // Worksheets

    /// <summary>
    /// Creates and returns a sheet after the last sheet in this workbook.
    /// </summary>
    worksheet create_sheet();

    /// <summary>
    /// Creates and returns a sheet at the specified index.
    /// </summary>
    worksheet create_sheet(std::size_t index);

    /// <summary>
    /// TODO: This should be private...
    /// </summary>
    worksheet create_sheet_with_rel(const std::string &title, const relationship &rel);

    /// <summary>
    /// Creates and returns a new sheet after the last sheet initializing it
    /// with all of the data from the provided worksheet.
    /// </summary>
    worksheet copy_sheet(worksheet worksheet);

    /// <summary>
    /// Creates and returns a new sheet at the specified index initializing it
    /// with all of the data from the provided worksheet.
    /// </summary>
    worksheet copy_sheet(worksheet worksheet, std::size_t index);

    /// <summary>
    /// Returns the worksheet that is determined to be active. An active
    /// sheet is that which is initially shown by the spreadsheet editor.
    /// </summary>
    worksheet active_sheet();

    /// <summary>
    /// Returns the worksheet with the given name. This may throw an exception
    /// if the sheet isn't found. Use workbook::contains(const std::string &)
    /// to make sure the sheet exists before calling this method.
    /// </summary>
    worksheet sheet_by_title(const std::string &title);

    /// <summary>
    /// Returns the worksheet with the given name. This may throw an exception
    /// if the sheet isn't found. Use workbook::contains(const std::string &)
    /// to make sure the sheet exists before calling this method.
    /// </summary>
    const worksheet sheet_by_title(const std::string &title) const;

    /// <summary>
    /// Returns the worksheet at the given index. This will throw an exception
    /// if index is greater than or equal to the number of sheets in this workbook.
    /// </summary>
    worksheet sheet_by_index(std::size_t index);

    /// <summary>
    /// Returns the worksheet at the given index. This will throw an exception
    /// if index is greater than or equal to the number of sheets in this workbook.
    /// </summary>
    const worksheet sheet_by_index(std::size_t index) const;

    /// <summary>
    /// Returns the worksheet with a sheetId of id. Sheet IDs are arbitrary numbers
    /// that uniquely identify a sheet. Most users won't need this.
    /// </summary>
    worksheet sheet_by_id(std::size_t id);

    /// <summary>
    /// Returns the worksheet with a sheetId of id. Sheet IDs are arbitrary numbers
    /// that uniquely identify a sheet. Most users won't need this.
    /// </summary>
    const worksheet sheet_by_id(std::size_t id) const;

    /// <summary>
    /// Returns true if this workbook contains a sheet with the given title.
    /// </summary>
    bool contains(const std::string &title) const;

    /// <summary>
    /// Returns the index of the given worksheet. The worksheet must be owned by this workbook.
    /// </summary>
    std::size_t index(worksheet worksheet);

    // remove worksheets

    /// <summary>
    /// Removes the given worksheet from this workbook.
    /// </summary>
    void remove_sheet(worksheet worksheet);

    /// <summary>
    /// Sets the contents of this workbook to be equivalent to that of
    /// a workbook returned by workbook::empty().
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
    /// Applies the function "f" to every non-empty cell in every worksheet in this workbook.
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

    // Metadata Properties

    /// <summary>
    /// Returns true if the workbook has the core property with the given name.
    /// </summary>
    bool has_core_property(xlnt::core_property type) const;

    /// <summary>
    /// Returns a vector of the type of each core property that is set to
    /// a particular value in this workbook.
    /// </summary>
    std::vector<xlnt::core_property> core_properties() const;

    /// <summary>
    /// Returns the value of the given core property.
    /// </summary>
    variant core_property(xlnt::core_property type) const;

    /// <summary>
    /// Sets the given core property to the provided value.
    /// </summary>
    void core_property(xlnt::core_property type, const variant &value);

    /// <summary>
    /// Returns true if the workbook has the extended property with the given name.
    /// </summary>
    bool has_extended_property(xlnt::extended_property type) const;

    /// <summary>
    /// Returns a vector of the type of each extended property that is set to
    /// a particular value in this workbook.
    /// </summary>
    std::vector<xlnt::extended_property> extended_properties() const;

    /// <summary>
    /// Returns the value of the given extended property.
    /// </summary>
    variant extended_property(xlnt::extended_property type) const;

    /// <summary>
    /// Sets the given extended property to the provided value.
    /// </summary>
    void extended_property(xlnt::extended_property type, const variant &value);

    /// <summary>
    /// Returns true if the workbook has the custom property with the given name.
    /// </summary>
    bool has_custom_property(const std::string &property_name) const;

    /// <summary>
    /// Returns a vector of the name of each custom property that is set to
    /// a particular value in this workbook.
    /// </summary>
    std::vector<std::string> custom_properties() const;

    /// <summary>
    /// Returns the value of the given custom property.
    /// </summary>
    variant custom_property(const std::string &property_name) const;

    /// <summary>
    /// Creates a new custom property in this workbook and sets it to the provided value.
    /// </summary>
    void custom_property(const std::string &property_name, const variant &value);

    /// <summary>
    /// Returns the base date used by this workbook. This will generally be windows_1900
    /// except on Apple based systems when it will default to mac_1904 unless otherwise
    /// set via `void workbook::base_date(calendar base_date)`.
    /// </summary>
    calendar base_date() const;

    /// <summary>
    /// Sets the base date style of this workbook. This is the date and time that
    /// a numeric value of 0 represents.
    /// </summary>
    void base_date(calendar base_date);

    /// <summary>
    /// Returns true if this workbook has had its title set.
    /// </summary>
    bool has_title() const;

    /// <summary>
    /// Returns the title of this workbook.
    /// </summary>
    std::string title() const;

    /// <summary>
    /// Sets the title of this workbook to title.
    /// </summary>
    void title(const std::string &title);

    /// <summary>
    /// Sets the absolute path of this workbook to path.
    /// </summary>
    void abs_path(const std::string &path);

    /// <summary>
    /// Sets the ArchID flags of this workbook to flags.
    /// </summary>
    void arch_id_flags(const std::size_t flags);

    // Named Ranges

    /// <summary>
    /// Returns a vector of the named ranges in this workbook.
    /// </summary>
    std::vector<xlnt::named_range> named_ranges() const;

    /// <summary>
    /// Creates a new names range.
    /// </summary>
    void create_named_range(const std::string &name, worksheet worksheet, const range_reference &reference);

    /// <summary>
    /// Creates a new names range.
    /// </summary>
    void create_named_range(const std::string &name, worksheet worksheet, const std::string &reference_string);

    /// <summary>
    /// Returns true if a named range of the given name exists in the workbook.
    /// </summary>
    bool has_named_range(const std::string &name) const;

    /// <summary>
    /// Returns the named range with the given name.
    /// </summary>
    class range named_range(const std::string &name);

    /// <summary>
    /// Deletes the named range with the given name.
    /// </summary>
    void remove_named_range(const std::string &name);

    // Serialization/Deserialization

    /// <summary>
    /// Serializes the workbook into an XLSX file and saves the bytes into
    /// byte vector data.
    /// </summary>
    void save(std::vector<std::uint8_t> &data) const;

    /// <summary>
    /// Serializes the workbook into an XLSX file encrypted with the given password
    /// and saves the bytes into byte vector data.
    /// </summary>
    void save(std::vector<std::uint8_t> &data, const std::string &password) const;

    /// <summary>
    /// Serializes the workbook into an XLSX file and saves the data into a file
    /// named filename.
    /// </summary>
    void save(const std::string &filename) const;

    /// <summary>
    /// Serializes the workbook into an XLSX file encrypted with the given password
    /// and loads the bytes into a file named filename.
    /// </summary>
    void save(const std::string &filename, const std::string &password) const;

#ifdef _MSC_VER
    /// <summary>
    /// Serializes the workbook into an XLSX file and saves the data into a file
    /// named filename.
    /// </summary>
    void save(const std::wstring &filename) const;

    /// <summary>
    /// Serializes the workbook into an XLSX file encrypted with the given password
    /// and loads the bytes into a file named filename.
    /// </summary>
    void save(const std::wstring &filename, const std::string &password) const;
#endif

    /// <summary>
    /// Serializes the workbook into an XLSX file and saves the data into a file
    /// named filename.
    /// </summary>
    void save(const xlnt::path &filename) const;

    /// <summary>
    /// Serializes the workbook into an XLSX file encrypted with the given password
    /// and loads the bytes into a file named filename.
    /// </summary>
    void save(const xlnt::path &filename, const std::string &password) const;

    /// <summary>
    /// Serializes the workbook into an XLSX file and saves the data into stream.
    /// </summary>
    void save(std::ostream &stream) const;

    /// <summary>
    /// Serializes the workbook into an XLSX file encrypted with the given password
    /// and loads the bytes into the given stream.
    /// </summary>
    void save(std::ostream &stream, const std::string &password) const;

    /// <summary>
    /// Interprets byte vector data as an XLSX file and sets the content of this
    /// workbook to match that file.
    /// </summary>
    void load(const std::vector<std::uint8_t> &data);

    /// <summary>
    /// Interprets byte vector data as an XLSX file encrypted with the
    /// given password and sets the content of this workbook to match that file.
    /// </summary>
    void load(const std::vector<std::uint8_t> &data, const std::string &password);

    /// <summary>
    /// Interprets file with the given filename as an XLSX file and sets
    /// the content of this workbook to match that file.
    /// </summary>
    void load(const std::string &filename);

    /// <summary>
    /// Interprets file with the given filename as an XLSX file encrypted with the
    /// given password and sets the content of this workbook to match that file.
    /// </summary>
    void load(const std::string &filename, const std::string &password);

#ifdef _MSC_VER
    /// <summary>
    /// Interprets file with the given filename as an XLSX file and sets
    /// the content of this workbook to match that file.
    /// </summary>
    void load(const std::wstring &filename);

    /// <summary>
    /// Interprets file with the given filename as an XLSX file encrypted with the
    /// given password and sets the content of this workbook to match that file.
    /// </summary>
    void load(const std::wstring &filename, const std::string &password);
#endif

    /// <summary>
    /// Interprets file with the given filename as an XLSX file and sets the
    /// content of this workbook to match that file.
    /// </summary>
    void load(const xlnt::path &filename);

    /// <summary>
    /// Interprets file with the given filename as an XLSX file encrypted with the
    /// given password and sets the content of this workbook to match that file.
    /// </summary>
    void load(const xlnt::path &filename, const std::string &password);

    /// <summary>
    /// Interprets data in stream as an XLSX file and sets the content of this
    /// workbook to match that file.
    /// </summary>
    void load(std::istream &stream);

    /// <summary>
    /// Interprets data in stream as an XLSX file encrypted with the given password
    /// and sets the content of this workbook to match that file.
    /// </summary>
    void load(std::istream &stream, const std::string &password);

    // View

    /// <summary>
    /// Returns true if this workbook has a view.
    /// </summary>
    bool has_view() const;

    /// <summary>
    /// Returns the view.
    /// </summary>
    workbook_view view() const;

    /// <summary>
    /// Sets the view to view.
    /// </summary>
    void view(const workbook_view &view);

    // Properties

    /// <summary>
    /// Returns true if a code name has been set for this workbook.
    /// </summary>
    bool has_code_name() const;

    /// <summary>
    /// Returns the code name that was set for this workbook.
    /// </summary>
    std::string code_name() const;

    /// <summary>
    /// Sets the code name of this workbook to code_name.
    /// </summary>
    void code_name(const std::string &code_name);

    /// <summary>
    /// Returns true if this workbook has a file version.
    /// </summary>
    bool has_file_version() const;

    /// <summary>
    /// Returns the AppName workbook file property.
    /// </summary>
    std::string app_name() const;

    /// <summary>
    /// Returns the LastEdited workbook file property.
    /// </summary>
    std::size_t last_edited() const;

    /// <summary>
    /// Returns the LowestEdited workbook file property.
    /// </summary>
    std::size_t lowest_edited() const;

    /// <summary>
    /// Returns the RupBuild workbook file property.
    /// </summary>
    std::size_t rup_build() const;

    // Theme

    /// <summary>
    /// Returns true if this workbook has a theme defined.
    /// </summary>
    bool has_theme() const;

    /// <summary>
    /// Returns a const reference to this workbook's theme.
    /// </summary>
    const xlnt::theme &theme() const;

    /// <summary>
    /// Sets the theme to value.
    /// </summary>
    void theme(const class theme &value);

    // Formats

    /// <summary>
    /// Returns the cell format at the given index. The index is the position of
    /// the format in xl/styles.xml.
    /// </summary>
    xlnt::format format(std::size_t format_index);

    /// <summary>
    /// Returns the cell format at the given index. The index is the position of
    /// the format in xl/styles.xml.
    /// </summary>
    const xlnt::format format(std::size_t format_index) const;

    /// <summary>
    /// Creates a new format and returns it.
    /// </summary>
    xlnt::format create_format(bool default_format = false);

    /// <summary>
    /// Clear all cell-level formatting and formats from the styelsheet. This leaves
    /// all other styling in place (e.g. named styles).
    /// </summary>
    void clear_formats();

    // Styles

    /// <summary>
    /// Returns true if this workbook has a style with a name of name.
    /// </summary>
    bool has_style(const std::string &name) const;

    /// <summary>
    /// Returns the named style with the given name.
    /// </summary>
    class style style(const std::string &name);

    /// <summary>
    /// Returns the named style with the given name.
    /// </summary>
    const class style style(const std::string &name) const;

    /// <summary>
    /// Creates a new style and returns it.
    /// </summary>
    class style create_style(const std::string &name);

    /// <summary>
    /// Creates a new style and returns it.
    /// </summary>
    class style create_builtin_style(std::size_t builtin_id);

    /// <summary>
    /// Clear all named styles from cells and remove the styles from
    /// from the styelsheet. This leaves all other styling in place
    /// (e.g. cell formats).
    /// </summary>
    void clear_styles();

    /// <summary>
    /// Sets the default slicer style to the given value.
    /// </summary>
    void default_slicer_style(const std::string &value);

    /// <summary>
    /// Returns the default slicer style.
    /// </summary>
    std::string default_slicer_style() const;

    /// <summary>
    /// Enables knownFonts in stylesheet.
    /// </summary>
    void enable_known_fonts();

    /// <summary>
    /// Disables knownFonts in stylesheet.
    /// </summary>
    void disable_known_fonts();

    /// <summary>
    /// Returns true if knownFonts are enabled in the stylesheet.
    /// </summary>
    bool known_fonts_enabled() const;

    // Manifest

    /// <summary>
    /// Returns a reference to the workbook's internal manifest.
    /// </summary>
    class manifest &manifest();

    /// <summary>
    /// Returns a reference to the workbook's internal manifest.
    /// </summary>
    const class manifest &manifest() const;

    // shared strings

    /// <summary>
    /// Append a shared string to the shared string collection in this workbook.
    /// This should not generally be called unless you know what you're doing.
    /// If allow_duplicates is false and the string is already in the collection,
    /// it will not be added. Returns the index of the added string.
    /// </summary>
    std::size_t add_shared_string(const rich_text &shared, bool allow_duplicates = false);

    /// <summary>
    /// Returns a reference to the shared string ordered by id
    /// </summary>
    const std::map<std::size_t, rich_text> &shared_strings_by_id() const;

    /// <summary>
    /// Returns a reference to the shared string related to the specified index
    /// </summary>
    const rich_text &shared_strings(std::size_t index) const;

    /// <summary>
    /// Returns a reference to the shared strings being used by cells
    /// in this workbook.
    /// </summary>
    std::unordered_map<rich_text, std::size_t, rich_text_hash> &shared_strings();

    /// <summary>
    /// Returns a reference to the shared strings being used by cells
    /// in this workbook.
    /// </summary>
    const std::unordered_map<rich_text, std::size_t, rich_text_hash> &shared_strings() const;

    // Thumbnail

    /// <summary>
    /// Sets the workbook's thumbnail to the given vector of bytes, thumbnail,
    /// with the given extension (e.g. jpg) and content_type (e.g. image/jpeg).
    /// </summary>
    void thumbnail(const std::vector<std::uint8_t> &thumbnail,
        const std::string &extension, const std::string &content_type);

    /// <summary>
    /// Returns a vector of bytes representing the workbook's thumbnail.
    /// </summary>
    const std::vector<std::uint8_t> &thumbnail() const;

    // Calculation properties

    /// <summary>
    /// Returns true if this workbook has any calculation properties set.
    /// </summary>
    bool has_calculation_properties() const;

    /// <summary>
    /// Returns the calculation properties used in this workbook.
    /// </summary>
    class calculation_properties calculation_properties() const;

    /// <summary>
    /// Sets the calculation properties of this workbook to props.
    /// </summary>
    void calculation_properties(const class calculation_properties &props);

    // Operators

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
    friend class streaming_workbook_reader;
    friend class worksheet;
    friend class detail::xlsx_consumer;
    friend class detail::xlsx_producer;

    /// <summary>
    /// Private constructor. Constructs a workbook from an implementation pointer.
    /// Used by static constructor to resolve circular construction dependency.
    /// </summary>
    workbook(detail::workbook_impl *impl);

    /// <summary>
    /// Returns a reference to the workbook implementation structure. Provides
    /// a nicer interface than constantly dereferencing workbook::d_.
    /// </summary>
    detail::workbook_impl &impl();

    /// <summary>
    /// Returns a reference to the workbook implementation structure. Provides
    /// a nicer interface than constantly dereferencing workbook::d_.
    /// </summary>
    const detail::workbook_impl &impl() const;

    /// <summary>
    /// Adds a package-level part of the given type to the manifest if it doesn't
    /// already exist. The part will have a path and content type of the default
    /// for that particular relationship type.
    /// </summary>
    void register_package_part(relationship_type type);

    /// <summary>
    /// Adds a workbook-level part of the given type to the manifest if it doesn't
    /// already exist. The part will have a path and content type of the default
    /// for that particular relationship type. It will be a relationship target
    /// of this workbook.
    /// </summary>
    void register_workbook_part(relationship_type type);

    /// <summary>
    /// Adds a worksheet-level part of the given type to the manifest if it doesn't
    /// already exist. The part will have a path and content type of the default
    /// for that particular relationship type. It will be a relationship target
    /// of the given worksheet, ws.
    /// </summary>
    void register_worksheet_part(worksheet ws, relationship_type type);

    /// <summary>
    /// Removes calcChain part from manifest if no formulae remain in workbook.
    /// </summary>
    void garbage_collect_formulae();

    /// <summary>
    /// Update extended workbook properties titlesOfParts and headingPairs when sheets change.
    /// </summary>
    void update_sheet_properties();

    /// <summary>
    /// Swaps the data held in this workbook with workbook other.
    /// </summary>
    void swap(workbook &other);

    /// <summary>
    /// Sheet 1 should be rId1, sheet 2 should be rId2, etc.
    /// </summary>
    void reorder_relationships();

    /// <summary>
    /// An opaque pointer to a structure that holds all of the data relating to this workbook.
    /// </summary>
    std::unique_ptr<detail::workbook_impl> d_;
};

} // namespace xlnt
