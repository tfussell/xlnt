
# API Reference

## Cell Module

### cell class

Describes cell associated properties.

Properties of interest include style, type, value, and address. The Cell class is required to know its value and type, display options, and any other features of an Excel cell.Utilities for referencing cells using Excel’s ‘A1’ column/row nomenclature are also provided.

#### `cell(const cell&)`
Default copy constructor.

#### `bool has_value() const`
Return true if value has been set and has not been cleared using cell::clear_value().

#### `template <typename T> T value() const`
Return the value of this cell as an instance of type T. Overloads exist for most C++ fundamental types like bool, int, etc. as well as for std::string and xlnt datetime types: date, time, datetime, and timedelta.

#### `void clear_value()`
Make this cell have a value of type null. All other cell attributes are retained.

#### `template <typename T> void value(T value)`
Set the value of this cell to the given value. Overloads exist for most C++ fundamental types like bool, int, etc. as well as for std::string and xlnt datetime types: date, time, datetime, and timedelta.

#### `void value(const std::string &string_value, bool infer_type)`
Analyze string_value to determine its type, convert it to that type, and set the value of this cell to that converted value.

#### `type data_type() const`
Return the type of this cell.

#### `void data_type(type t)`
Set the type of this cell.

#### `bool garbage_collectible() const`
There’s no reason to keep a cell which has no value and is not a placeholder. Return true if this cell has no value, style, isn’t merged, etc.

#### `bool is_date() const`
Return true iff this cell’s number format matches a date format.

#### `cell_reference reference() const`
Return a cell_reference that points to the location of this cell.

#### `column_t column() const`
Return the column of this cell.

#### `row_t row() const`
Return the row of this cell.

#### `std::pair<int, int> anchor() const`
Return the location of this cell as an ordered pair.

#### `std::string hyperlink() const`
Return the URL of this cell’s hyperlink.

#### `void hyperlink(const std::string &value)`
Add a hyperlink to this cell pointing to the URI of the given value.

#### `bool has_hyperlink() const`
Return true if this cell has a hyperlink set.

#### `class alignment computed_alignment() const`
Returns the result of computed_format().alignment().

#### `class border computed_border() const`
Returns the result of computed_format().border().

#### `class fill computed_fill() const`
Returns the result of computed_format().fill().

#### `class font computed_font() const`
Returns the result of computed_format().font().

#### `class number_format computed_number_format() const`
Returns the result of computed_format().number_format().

#### `class protection computed_protection() const`
Returns the result of computed_format().protection().

#### `bool has_format() const`
Return true if this cell has had a format applied to it.

#### `const class format format() const`
Return a reference to the format applied to this cell. If this cell has no format, an invalid_attribute exception will be thrown.

#### `void format(const class format new_format)`
Applies the cell-level formatting of new_format to this cell.

#### `void clear_format()`
Remove the cell-level formatting from this cell. This doesn’t affect the style that may also be applied to the cell. Throws an invalid_attribute exception if no format is applied.

#### `class number_format number_format() const`
Returns the number format of this cell.

#### `void number_format(const class number_format &format)`
Creates a new format in the workbook, sets its number_format to the given format, and applies the format to this cell.

#### `class font font() const`
Returns the font applied to the text in this cell.

#### `void font(const class font &font_)`
Creates a new format in the workbook, sets its font to the given font, and applies the format to this cell.

#### `class fill fill() const`
Returns the fill applied to this cell.

#### `void fill(const class fill &fill_)`
Creates a new format in the workbook, sets its fill to the given fill, and applies the format to this cell.

#### `class border border() const`
Returns the border of this cell.

#### `void border(const class border &border_)`
Creates a new format in the workbook, sets its border to the given border, and applies the format to this cell.

#### `class alignment alignment() const`
Returns the alignment of the text in this cell.

#### `void alignment(const class alignment &alignment_)`
Creates a new format in the workbook, sets its alignment to the given alignment, and applies the format to this cell.

#### `class protection protection() const`
Returns the protection of this cell.

#### `void protection(const class protection &protection_)`
Creates a new format in the workbook, sets its protection to the given protection, and applies the format to this cell.

#### `bool has_style() const`
Returns true if this cell has had a style applied to it.

#### `const class style style() const`
Returns a wrapper pointing to the named style applied to this cell.

#### `void style(const class style &new_style)`
Equivalent to style(new_style.name())

#### `void style(const std::string &style_name)`
Sets the named style applied to this cell to a style named style_name. If this style has not been previously created in the workbook, a key_not_found exception will be thrown.

#### `void clear_style()`
Removes the named style from this cell. An invalid_attribute exception will be thrown if this cell has no style. This will not affect the cell format of the cell.

#### `std::string formula() const`
Returns the string representation of the formula applied to this cell.

#### `void formula(const std::string &formula)`
Sets the formula of this cell to the given value. This formula string should begin with ‘=’.

#### `void clear_formula()`
Removes the formula from this cell. After this is called, has_formula() will return false.

#### `bool has_formula() const`
Returns true if this cell has had a formula applied to it.

#### `std::string to_string() const`
Returns a string representing the value of this cell. If the data type is not a string, it will be converted according to the number format.

#### `bool is_merged() const`
Return true iff this cell has been merged with one or more surrounding cells.

#### `void merged(bool merged)`
Make this a merged cell iff merged is true. Generally, this shouldn’t be called directly. Instead, use worksheet::merge_cells on its parent worksheet.

#### `std::string error() const`
Return the error string that is stored in this cell.

#### `void error(const std::string &error)`
Directly assign the value of this cell to be the given error.

#### `cell offset(int column, int row)`
Return a cell from this cell’s parent workbook at a relative offset given by the parameters.

#### `class worksheet worksheet()`
Return the worksheet that owns this cell.

#### `const class worksheet worksheet() const`
Return the worksheet that owns this cell.

#### `class workbook &workbook()`
Return the workbook of the worksheet that owns this cell.

#### `const class workbook &workbook() const`
Return the workbook of the worksheet that owns this cell.

#### `calendar base_date() const`
Shortcut to return the base date of the parent workbook. Equivalent to workbook().properties().excel_base_date

#### `std::string check_string(const std::string &to_check)`
Return to_check after checking encoding, size, and illegal characters.

#### `bool has_comment()`
Return true if this cell has a comment applied.

#### `void clear_comment()`
Delete the comment applied to this cell if it exists.

#### `class comment comment()`
Get the comment applied to this cell.

#### `void comment(const std::string &text, const std::string &author = "Microsoft Office User")`
Create a new comment with the given text and optional author and apply it to the cell.

#### `void comment(const std::string &comment_text, const class font &comment_font, const std::string &author = "Microsoft Office User")`
Create a new comment with the given text, formatting, and optional author and apply it to the cell.

#### `void comment(const class comment &new_comment)`
Apply the comment provided as the only argument to the cell.

#### `cell &operator=(const cell &rhs)`
Make this cell point to rhs. The cell originally pointed to by this cell will be unchanged.

#### `bool operator==(const cell &comparand) const`
Return true if this cell the same cell as comparand (compare by reference).

#### `bool operator==(std::nullptr_t) const`
Return true if this cell is uninitialized.

#### `static const std::unordered_map<std::string, int> &error_codes()`
Return a map of error strings such as #DIV/0! and their associated indices.

### cell_reference class

#### `cell_reference()`

Default constructor makes a reference to the top-left-most cell, “A1”.

#### `cell_reference(const char *reference_string)`

Constructs a cell_reference from a string reprenting a cell coordinate (e.g. $B14).

## Workbook Module

### workbook class

#### `workbook::workbook()`

Constructs a workbook.
