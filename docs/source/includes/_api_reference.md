# API Reference
## Cell Module
### cell
#### ```using xlnt::cell::type =  cell_typeundefined```
Alias xlnt::cell_type to xlnt::cell::type since it looks nicer.
#### ```friend class detail::xlsx_consumerundefined```
#### ```friend class detail::xlsx_producerundefined```
#### ```friend struct detail::cell_implundefined```
#### ```static const std::unordered_map<std::string, int>& xlnt::cell::error_codes()```
Returns a map of error strings such as #DIV/0! and their associated indices.
#### ```xlnt::cell::cell(const cell &)=default```
Default copy constructor.
#### ```bool xlnt::cell::has_value() const```
Returns true if value has been set and has not been cleared using cell::clear_value().
#### ```T xlnt::cell::value() const```
Returns the value of this cell as an instance of type T. Overloads exist for most C++ fundamental types like bool, int, etc. as well as for std::string and xlnt datetime types: date, time, datetime, and timedelta.
#### ```void xlnt::cell::clear_value()```
Makes this cell have a value of type null. All other cell attributes are retained.
#### ```void xlnt::cell::value(std::nullptr_t)```
Sets the type of this cell to null.
#### ```void xlnt::cell::value(bool boolean_value)```
Sets the value of this cell to the given boolean value.
#### ```void xlnt::cell::value(int int_value)```
Sets the value of this cell to the given value.
#### ```void xlnt::cell::value(unsigned int int_value)```
Sets the value of this cell to the given value.
#### ```void xlnt::cell::value(long long int int_value)```
Sets the value of this cell to the given value.
#### ```void xlnt::cell::value(unsigned long long int int_value)```
Sets the value of this cell to the given value.
#### ```void xlnt::cell::value(float float_value)```
Sets the value of this cell to the given value.
#### ```void xlnt::cell::value(double float_value)```
Sets the value of this cell to the given value.
#### ```void xlnt::cell::value(long double float_value)```
Sets the value of this cell to the given value.
#### ```void xlnt::cell::value(const date &date_value)```
Sets the value of this cell to the given value.
#### ```void xlnt::cell::value(const time &time_value)```
Sets the value of this cell to the given value.
#### ```void xlnt::cell::value(const datetime &datetime_value)```
Sets the value of this cell to the given value.
#### ```void xlnt::cell::value(const timedelta &timedelta_value)```
Sets the value of this cell to the given value.
#### ```void xlnt::cell::value(const std::string &string_value)```
Sets the value of this cell to the given value.
#### ```void xlnt::cell::value(const char *string_value)```
Sets the value of this cell to the given value.
#### ```void xlnt::cell::value(const rich_text &text_value)```
Sets the value of this cell to the given value.
#### ```void xlnt::cell::value(const cell other_cell)```
Sets the value and formatting of this cell to that of other_cell.
#### ```void xlnt::cell::value(const std::string &string_value, bool infer_type)```
Analyzes string_value to determine its type, convert it to that type, and set the value of this cell to that converted value.
#### ```type xlnt::cell::data_type() const```
Returns the type of this cell.
#### ```void xlnt::cell::data_type(type t)```
Sets the type of this cell. This should usually be done indirectly by setting the value of the cell to a value of that type.
#### ```bool xlnt::cell::garbage_collectible() const```
There's no reason to keep a cell which has no value and is not a placeholder. Returns true if this cell has no value, style, isn't merged, etc.
#### ```bool xlnt::cell::is_date() const```
Returns true iff this cell's number format matches a date format.
#### ```cell_reference xlnt::cell::reference() const```
Returns a cell_reference that points to the location of this cell.
#### ```column_t xlnt::cell::column() const```
Returns the column of this cell.
#### ```row_t xlnt::cell::row() const```
Returns the row of this cell.
#### ```std::pair<int, int> xlnt::cell::anchor() const```
Returns the location of this cell as an ordered pair (left, top).
#### ```std::string xlnt::cell::hyperlink() const```
Returns the URL of this cell's hyperlink.
#### ```void xlnt::cell::hyperlink(const std::string &url)```
Adds a hyperlink to this cell pointing to the URL of the given value.
#### ```void xlnt::cell::hyperlink(const std::string &url, const std::string &display)```
Adds a hyperlink to this cell pointing to the URI of the given value and sets the text value of the cell to the given parameter.
#### ```void xlnt::cell::hyperlink(xlnt::cell target)```
Adds an internal hyperlink to this cell pointing to the given cell.
#### ```bool xlnt::cell::has_hyperlink() const```
Returns true if this cell has a hyperlink set.
#### ```class alignment xlnt::cell::computed_alignment() const```
Returns the alignment that should be used when displaying this cell graphically based on the workbook default, the cell-level format, and the named style applied to the cell in that order.
#### ```class border xlnt::cell::computed_border() const```
Returns the border that should be used when displaying this cell graphically based on the workbook default, the cell-level format, and the named style applied to the cell in that order.
#### ```class fill xlnt::cell::computed_fill() const```
Returns the fill that should be used when displaying this cell graphically based on the workbook default, the cell-level format, and the named style applied to the cell in that order.
#### ```class font xlnt::cell::computed_font() const```
Returns the font that should be used when displaying this cell graphically based on the workbook default, the cell-level format, and the named style applied to the cell in that order.
#### ```class number_format xlnt::cell::computed_number_format() const```
Returns the number format that should be used when displaying this cell graphically based on the workbook default, the cell-level format, and the named style applied to the cell in that order.
#### ```class protection xlnt::cell::computed_protection() const```
Returns the protection that should be used when displaying this cell graphically based on the workbook default, the cell-level format, and the named style applied to the cell in that order.
#### ```bool xlnt::cell::has_format() const```
Returns true if this cell has had a format applied to it.
#### ```const class format xlnt::cell::format() const```
Returns the format applied to this cell. If this cell has no format, an invalid_attribute exception will be thrown.
#### ```void xlnt::cell::format(const class format new_format)```
Applies the cell-level formatting of new_format to this cell.
#### ```void xlnt::cell::clear_format()```
Removes the cell-level formatting from this cell. This doesn't affect the style that may also be applied to the cell. Throws an invalid_attribute exception if no format is applied.
#### ```class number_format xlnt::cell::number_format() const```
Returns the number format of this cell.
#### ```void xlnt::cell::number_format(const class number_format &format)```
Creates a new format in the workbook, sets its number_format to the given format, and applies the format to this cell.
#### ```class font xlnt::cell::font() const```
Returns the font applied to the text in this cell.
#### ```void xlnt::cell::font(const class font &font_)```
Creates a new format in the workbook, sets its font to the given font, and applies the format to this cell.
#### ```class fill xlnt::cell::fill() const```
Returns the fill applied to this cell.
#### ```void xlnt::cell::fill(const class fill &fill_)```
Creates a new format in the workbook, sets its fill to the given fill, and applies the format to this cell.
#### ```class border xlnt::cell::border() const```
Returns the border of this cell.
#### ```void xlnt::cell::border(const class border &border_)```
Creates a new format in the workbook, sets its border to the given border, and applies the format to this cell.
#### ```class alignment xlnt::cell::alignment() const```
Returns the alignment of the text in this cell.
#### ```void xlnt::cell::alignment(const class alignment &alignment_)```
Creates a new format in the workbook, sets its alignment to the given alignment, and applies the format to this cell.
#### ```class protection xlnt::cell::protection() const```
Returns the protection of this cell.
#### ```void xlnt::cell::protection(const class protection &protection_)```
Creates a new format in the workbook, sets its protection to the given protection, and applies the format to this cell.
#### ```bool xlnt::cell::has_style() const```
Returns true if this cell has had a style applied to it.
#### ```class style xlnt::cell::style()```
Returns a wrapper pointing to the named style applied to this cell.
#### ```const class style xlnt::cell::style() const```
Returns a wrapper pointing to the named style applied to this cell.
#### ```void xlnt::cell::style(const class style &new_style)```
Sets the named style applied to this cell to a style named style_name. Equivalent to style(new_style.name()).
#### ```void xlnt::cell::style(const std::string &style_name)```
Sets the named style applied to this cell to a style named style_name. If this style has not been previously created in the workbook, a key_not_found exception will be thrown.
#### ```void xlnt::cell::clear_style()```
Removes the named style from this cell. An invalid_attribute exception will be thrown if this cell has no style. This will not affect the cell format of the cell.
#### ```std::string xlnt::cell::formula() const```
Returns the string representation of the formula applied to this cell.
#### ```void xlnt::cell::formula(const std::string &formula)```
Sets the formula of this cell to the given value. This formula string should begin with '='.
#### ```void xlnt::cell::clear_formula()```
Removes the formula from this cell. After this is called, has_formula() will return false.
#### ```bool xlnt::cell::has_formula() const```
Returns true if this cell has had a formula applied to it.
#### ```std::string xlnt::cell::to_string() const```
Returns a string representing the value of this cell. If the data type is not a string, it will be converted according to the number format.
#### ```bool xlnt::cell::is_merged() const```
Returns true iff this cell has been merged with one or more surrounding cells.
#### ```void xlnt::cell::merged(bool merged)```
Makes this a merged cell iff merged is true. Generally, this shouldn't be called directly. Instead, use worksheet::merge_cells on its parent worksheet.
#### ```std::string xlnt::cell::error() const```
Returns the error string that is stored in this cell.
#### ```void xlnt::cell::error(const std::string &error)```
Directly assigns the value of this cell to be the given error.
#### ```cell xlnt::cell::offset(int column, int row)```
Returns a cell from this cell's parent workbook at a relative offset given by the parameters.
#### ```class worksheet xlnt::cell::worksheet()```
Returns the worksheet that owns this cell.
#### ```const class worksheet xlnt::cell::worksheet() const```
Returns the worksheet that owns this cell.
#### ```class workbook& xlnt::cell::workbook()```
Returns the workbook of the worksheet that owns this cell.
#### ```const class workbook& xlnt::cell::workbook() const```
Returns the workbook of the worksheet that owns this cell.
#### ```calendar xlnt::cell::base_date() const```
Returns the base date of the parent workbook.
#### ```std::string xlnt::cell::check_string(const std::string &to_check)```
Returns to_check after verifying and fixing encoding, size, and illegal characters.
#### ```bool xlnt::cell::has_comment()```
Returns true if this cell has a comment applied.
#### ```void xlnt::cell::clear_comment()```
Deletes the comment applied to this cell if it exists.
#### ```class comment xlnt::cell::comment()```
Gets the comment applied to this cell.
#### ```void xlnt::cell::comment(const std::string &text, const std::string &author="Microsoft Office User")```
Creates a new comment with the given text and optional author and applies it to the cell.
#### ```void xlnt::cell::comment(const std::string &comment_text, const class font &comment_font, const std::string &author="Microsoft Office User")```
Creates a new comment with the given text, formatting, and optional author and applies it to the cell.
#### ```void xlnt::cell::comment(const class comment &new_comment)```
Apply the comment provided as the only argument to the cell.
#### ```double xlnt::cell::width() const```
Returns the width of this cell in pixels.
#### ```double xlnt::cell::height() const```
Returns the height of this cell in pixels.
#### ```cell& xlnt::cell::operator=(const cell &rhs)```
Makes this cell interally point to rhs. The cell data originally pointed to by this cell will be unchanged.
#### ```bool xlnt::cell::operator==(const cell &comparand) const```
Returns true if this cell the same cell as comparand (compared by reference).
#### ```bool xlnt::cell::operator==(std::nullptr_t) const```
Returns true if this cell is uninitialized.
### cell_reference_hash
#### ```std::size_t xlnt::cell_reference_hash::operator()(const cell_reference &k) const```
Returns a hash representing a particular cell_reference.
### cell_reference
#### ```static std::pair<std::string, row_t> xlnt::cell_reference::split_reference(const std::string &reference_string)```
Splits a coordinate string like "A1" into an equivalent pair like {"A", 1}.
#### ```static std::pair<std::string, row_t> xlnt::cell_reference::split_reference(const std::string &reference_string, bool &absolute_column, bool &absolute_row)```
Splits a coordinate string like "A1" into an equivalent pair like {"A", 1}. Reference parameters absolute_column and absolute_row will be set to true if column part or row part are prefixed by a dollar-sign indicating they are absolute, otherwise false.
#### ```xlnt::cell_reference::cell_reference()```
Default constructor makes a reference to the top-left-most cell, "A1".
#### ```xlnt::cell_reference::cell_reference(const char *reference_string)```
Constructs a cell_reference from a string reprenting a cell coordinate (e.g. $B14).
#### ```xlnt::cell_reference::cell_reference(const std::string &reference_string)```
Constructs a cell_reference from a string reprenting a cell coordinate (e.g. $B14).
#### ```xlnt::cell_reference::cell_reference(column_t column, row_t row)```
Constructs a cell_reference from a 1-indexed column index and row index.
#### ```cell_reference& xlnt::cell_reference::make_absolute(bool absolute_column=true, bool absolute_row=true)```
Converts a coordinate to an absolute coordinate string (e.g. B12 -> $B$12) Defaulting to true, absolute_column and absolute_row can optionally control whether the resulting cell_reference has an absolute column (e.g. B12 -> $B12) and absolute row (e.g. B12 -> B$12) respectively.
#### ```bool xlnt::cell_reference::column_absolute() const```
Returns true if the reference refers to an absolute column, otherwise false.
#### ```void xlnt::cell_reference::column_absolute(bool absolute_column)```
Makes this reference have an absolute column if absolute_column is true, otherwise not absolute.
#### ```bool xlnt::cell_reference::row_absolute() const```
Returns true if the reference refers to an absolute row, otherwise false.
#### ```void xlnt::cell_reference::row_absolute(bool absolute_row)```
Makes this reference have an absolute row if absolute_row is true, otherwise not absolute.
#### ```column_t xlnt::cell_reference::column() const```
Returns a string that identifies the column of this reference (e.g. second column from left is "B")
#### ```void xlnt::cell_reference::column(const std::string &column_string)```
Sets the column of this reference from a string that identifies a particular column.
#### ```column_t::index_t xlnt::cell_reference::column_index() const```
Returns a 1-indexed numeric index of the column of this reference.
#### ```void xlnt::cell_reference::column_index(column_t column)```
Sets the column of this reference from a 1-indexed number that identifies a particular column.
#### ```row_t xlnt::cell_reference::row() const```
Returns a 1-indexed numeric index of the row of this reference.
#### ```void xlnt::cell_reference::row(row_t row)```
Sets the row of this reference from a 1-indexed number that identifies a particular row.
#### ```cell_reference xlnt::cell_reference::make_offset(int column_offset, int row_offset) const```
Returns a cell_reference offset from this cell_reference by the number of columns and rows specified by the parameters. A negative value for column_offset or row_offset results in a reference above or left of this cell_reference, respectively.
#### ```std::string xlnt::cell_reference::to_string() const```
Returns a string like "A1" for cell_reference(1, 1).
#### ```range_reference xlnt::cell_reference::to_range() const```
Returns a 1x1 range_reference containing only this cell_reference.
#### ```range_reference xlnt::cell_reference::operator,(const cell_reference &other) const```
I've always wanted to overload the comma operator. cell_reference("A", 1), cell_reference("B", 1) will return range_reference(cell_reference("A", 1), cell_reference("B", 1))
#### ```bool xlnt::cell_reference::operator==(const cell_reference &comparand) const```
Returns true if this reference is identical to comparand including in absoluteness of column and row.
#### ```bool xlnt::cell_reference::operator==(const std::string &reference_string) const```
Constructs a cell_reference from reference_string and return the result of their comparison.
#### ```bool xlnt::cell_reference::operator==(const char *reference_string) const```
Constructs a cell_reference from reference_string and return the result of their comparison.
#### ```bool xlnt::cell_reference::operator!=(const cell_reference &comparand) const```
Returns true if this reference is not identical to comparand including in absoluteness of column and row.
#### ```bool xlnt::cell_reference::operator!=(const std::string &reference_string) const```
Constructs a cell_reference from reference_string and return the result of their comparison.
#### ```bool xlnt::cell_reference::operator!=(const char *reference_string) const```
Constructs a cell_reference from reference_string and return the result of their comparison.
### comment
#### ```xlnt::comment::comment()```
Constructs a new blank comment.
#### ```xlnt::comment::comment(const rich_text &text, const std::string &author)```
Constructs a new comment with the given text and author.
#### ```xlnt::comment::comment(const std::string &text, const std::string &author)```
Constructs a new comment with the given unformatted text and author.
#### ```rich_text xlnt::comment::text() const```
Returns the text that will be displayed for this comment.
#### ```std::string xlnt::comment::plain_text() const```
Returns the plain text that will be displayed for this comment without formatting information.
#### ```std::string xlnt::comment::author() const```
Returns the author of this comment.
#### ```void xlnt::comment::hide()```
Makes this comment only visible when the associated cell is hovered.
#### ```void xlnt::comment::show()```
Makes this comment always visible.
#### ```bool xlnt::comment::visible() const```
Returns true if this comment is not hidden.
#### ```void xlnt::comment::position(int left, int top)```
Sets the absolute position of this cell to the given coordinates.
#### ```int xlnt::comment::left() const```
Returns the distance from the left side of the sheet to the left side of the comment.
#### ```int xlnt::comment::top() const```
Returns the distance from the top of the sheet to the top of the comment.
#### ```void xlnt::comment::size(int width, int height)```
Sets the size of the comment.
#### ```int xlnt::comment::width() const```
Returns the width of this comment.
#### ```int xlnt::comment::height() const```
Returns the height of this comment.
#### ```bool xlnt::comment::operator==(const comment &other) const```
Return true if this comment is equivalent to other.
#### ```bool xlnt::comment::operator!=(const comment &other) const```
Returns true if this comment is not equivalent to other.
### column_t
#### ```using xlnt::column_t::index_t =  std::uint32_tundefined```
Alias declaration for the internal numeric type of this column.
#### ```index_t xlnt::column_t::indexundefined```
Internal numeric value of this column index.
#### ```static index_t xlnt::column_t::column_index_from_string(const std::string &column_string)```
Convert a column letter into a column number (e.g. B -> 2)
#### ```static std::string xlnt::column_t::column_string_from_index(index_t column_index)```
Convert a column number into a column letter (3 -> 'C')
#### ```xlnt::column_t::column_t()```
Default constructor. The column points to the "A" column.
#### ```xlnt::column_t::column_t(index_t column_index)```
Constructs a column from a number.
#### ```xlnt::column_t::column_t(const std::string &column_string)```
Constructs a column from a string.
#### ```xlnt::column_t::column_t(const char *column_string)```
Constructs a column from a string.
#### ```xlnt::column_t::column_t(const column_t &other)```
Copy constructor. Constructs a column by copying it from other.
#### ```xlnt::column_t::column_t(column_t &&other)```
Move constructor. Constructs a column by moving it from other.
#### ```std::string xlnt::column_t::column_string() const```
Returns a string representation of this column index.
#### ```column_t& xlnt::column_t::operator=(column_t rhs)```
Sets this column to be the same as rhs's and return reference to self.
#### ```column_t& xlnt::column_t::operator=(const std::string &rhs)```
Sets this column to be equal to rhs and return reference to self.
#### ```column_t& xlnt::column_t::operator=(const char *rhs)```
Sets this column to be equal to rhs and return reference to self.
#### ```bool xlnt::column_t::operator==(const column_t &other) const```
Returns true if this column refers to the same column as other.
#### ```bool xlnt::column_t::operator!=(const column_t &other) const```
Returns true if this column doesn't refer to the same column as other.
#### ```bool xlnt::column_t::operator==(int other) const```
Returns true if this column refers to the same column as other.
#### ```bool xlnt::column_t::operator==(index_t other) const```
Returns true if this column refers to the same column as other.
#### ```bool xlnt::column_t::operator==(const std::string &other) const```
Returns true if this column refers to the same column as other.
#### ```bool xlnt::column_t::operator==(const char *other) const```
Returns true if this column refers to the same column as other.
#### ```bool xlnt::column_t::operator!=(int other) const```
Returns true if this column doesn't refer to the same column as other.
#### ```bool xlnt::column_t::operator!=(index_t other) const```
Returns true if this column doesn't refer to the same column as other.
#### ```bool xlnt::column_t::operator!=(const std::string &other) const```
Returns true if this column doesn't refer to the same column as other.
#### ```bool xlnt::column_t::operator!=(const char *other) const```
Returns true if this column doesn't refer to the same column as other.
#### ```bool xlnt::column_t::operator>(const column_t &other) const```
Returns true if other is to the right of this column.
#### ```bool xlnt::column_t::operator>=(const column_t &other) const```
Returns true if other is to the right of or equal to this column.
#### ```bool xlnt::column_t::operator<(const column_t &other) const```
Returns true if other is to the left of this column.
#### ```bool xlnt::column_t::operator<=(const column_t &other) const```
Returns true if other is to the left of or equal to this column.
#### ```bool xlnt::column_t::operator>(const column_t::index_t &other) const```
Returns true if other is to the right of this column.
#### ```bool xlnt::column_t::operator>=(const column_t::index_t &other) const```
Returns true if other is to the right of or equal to this column.
#### ```bool xlnt::column_t::operator<(const column_t::index_t &other) const```
Returns true if other is to the left of this column.
#### ```bool xlnt::column_t::operator<=(const column_t::index_t &other) const```
Returns true if other is to the left of or equal to this column.
#### ```column_t& xlnt::column_t::operator++()```
Pre-increments this column, making it point to the column one to the right and returning a reference to it.
#### ```column_t& xlnt::column_t::operator--()```
Pre-deccrements this column, making it point to the column one to the left and returning a reference to it.
#### ```column_t xlnt::column_t::operator++(int)```
Post-increments this column, making it point to the column one to the right and returning the old column.
#### ```column_t xlnt::column_t::operator--(int)```
Post-decrements this column, making it point to the column one to the left and returning the old column.
#### ```column_t& xlnt::column_t::operator+=(const column_t &rhs)```
Adds rhs to this column and returns a reference to this column.
#### ```column_t& xlnt::column_t::operator-=(const column_t &rhs)```
Subtracts rhs from this column and returns a reference to this column.
#### ```column_t operator+(column_t lhs, const column_t &rhs)```
Returns the result of adding rhs to this column.
#### ```column_t operator-(column_t lhs, const column_t &rhs)```
Returns the result of subtracing lhs by rhs column.
#### ```bool operator>(const column_t::index_t &left, const column_t &right)```
Returns true if other is to the right of this column.
#### ```bool operator>=(const column_t::index_t &left, const column_t &right)```
Returns true if other is to the right of or equal to this column.
#### ```bool operator<(const column_t::index_t &left, const column_t &right)```
Returns true if other is to the left of this column.
#### ```bool operator<=(const column_t::index_t &left, const column_t &right)```
Returns true if other is to the left of or equal to this column.
#### ```void swap(column_t &left, column_t &right)```
Swaps the columns that left and right refer to.
### column_hash
#### ```std::size_t xlnt::column_hash::operator()(const column_t &k) const```
Returns the result of hashing column k.
### column_t >
#### ```size_t std::hash< xlnt::column_t >::operator()(const xlnt::column_t &k) const```
Returns the result of hashing column k.
### rich_text
#### ```xlnt::rich_text::rich_text()=default```
Constructs an empty rich text object with no font and empty text.
#### ```xlnt::rich_text::rich_text(const std::string &plain_text)```
Constructs a rich text object with the given text and no font.
#### ```xlnt::rich_text::rich_text(const std::string &plain_text, const class font &text_font)```
Constructs a rich text object with the given text and font.
#### ```xlnt::rich_text::rich_text(const rich_text_run &single_run)```
Copy constructor.
#### ```void xlnt::rich_text::clear()```
Removes all text runs from this text.
#### ```void xlnt::rich_text::plain_text(const std::string &s)```
Clears any runs in this text and adds a single run with default formatting and the given string as its textual content.
#### ```std::string xlnt::rich_text::plain_text() const```
Combines the textual content of each text run in order and returns the result.
#### ```std::vector<rich_text_run> xlnt::rich_text::runs() const```
Returns a copy of the individual runs that comprise this text.
#### ```void xlnt::rich_text::runs(const std::vector< rich_text_run > &new_runs)```
Sets the runs of this text all at once.
#### ```void xlnt::rich_text::add_run(const rich_text_run &t)```
Adds a new run to the end of the set of runs.
#### ```bool xlnt::rich_text::operator==(const rich_text &rhs) const```
Returns true if the runs that make up this text are identical to those in rhs.
#### ```bool xlnt::rich_text::operator==(const std::string &rhs) const```
Returns true if this text has a single unformatted run with text equal to rhs.
## Packaging Module
### manifest
#### ```void xlnt::manifest::clear()```
Unregisters all default and override type and all relationships and known parts.
#### ```std::vector<path> xlnt::manifest::parts() const```
Returns the path to all internal package parts registered as a source or target of a relationship.
#### ```bool xlnt::manifest::has_relationship(const path &source, relationship_type type) const```
Returns true if the manifest contains a relationship with the given type with part as the source.
#### ```bool xlnt::manifest::has_relationship(const path &source, const std::string &rel_id) const```
Returns true if the manifest contains a relationship with the given type with part as the source.
#### ```class relationship xlnt::manifest::relationship(const path &source, relationship_type type) const```
Returns the relationship with "source" as the source and with a type of "type". Throws a key_not_found exception if no such relationship is found.
#### ```class relationship xlnt::manifest::relationship(const path &source, const std::string &rel_id) const```
Returns the relationship with "source" as the source and with an ID of "rel_id". Throws a key_not_found exception if no such relationship is found.
#### ```std::vector<xlnt::relationship> xlnt::manifest::relationships(const path &source) const```
Returns all relationship with "source" as the source.
#### ```std::vector<xlnt::relationship> xlnt::manifest::relationships(const path &source, relationship_type type) const```
Returns all relationships with "source" as the source and with a type of "type".
#### ```path xlnt::manifest::canonicalize(const std::vector< xlnt::relationship > &rels) const```
Returns the canonical path of the chain of relationships by traversing through rels and forming the absolute combined path.
#### ```std::string xlnt::manifest::register_relationship(const uri &source, relationship_type type, const uri &target, target_mode mode)```
Registers a new relationship by specifying all of the relationship properties explicitly.
#### ```std::string xlnt::manifest::register_relationship(const class relationship &rel)```
Registers a new relationship already constructed elsewhere.
#### ```std::unordered_map<std::string, std::string> xlnt::manifest::unregister_relationship(const uri &source, const std::string &rel_id)```
Delete the relationship with the given id from source part. Returns a mapping of relationship IDs since IDs are shifted down. For example, if there are three relationships for part a.xml like [rId1, rId2, rId3] and rId2 is deleted, the resulting map would look like [rId3->rId2].
#### ```std::string xlnt::manifest::content_type(const path &part) const```
Given the path to a part, returns the content type of the part as a string.
#### ```bool xlnt::manifest::has_default_type(const std::string &extension) const```
Returns true if a default content type for the extension has been registered.
#### ```std::vector<std::string> xlnt::manifest::extensions_with_default_types() const```
Returns a vector of all extensions with registered default content types.
#### ```std::string xlnt::manifest::default_type(const std::string &extension) const```
Returns the registered default content type for parts of the given extension.
#### ```void xlnt::manifest::register_default_type(const std::string &extension, const std::string &type)```
Associates the given extension with the given content type.
#### ```void xlnt::manifest::unregister_default_type(const std::string &extension)```
Unregisters the default content type for the given extension.
#### ```bool xlnt::manifest::has_override_type(const path &part) const```
Returns true if a content type overriding the default content type has been registered for the given part.
#### ```std::string xlnt::manifest::override_type(const path &part) const```
Returns the override content type registered for the given part. Throws key_not_found exception if no override type was found.
#### ```std::vector<path> xlnt::manifest::parts_with_overriden_types() const```
Returns the path of every part in this manifest with an overriden content type.
#### ```void xlnt::manifest::register_override_type(const path &part, const std::string &type)```
Overrides any default type registered for the part's extension with the given content type.
#### ```void xlnt::manifest::unregister_override_type(const path &part)```
Unregisters the overriding content type of the given part.
### relationship
#### ```xlnt::relationship::relationship()```
Constructs a new empty relationship.
#### ```xlnt::relationship::relationship(const std::string &id, relationship_type t, const uri &source, const uri &target, xlnt::target_mode mode)```
Constructs a new relationship by specifying all of its properties.
#### ```std::string xlnt::relationship::id() const```
Returns a string of the form rId# that identifies the relationship.
#### ```relationship_type xlnt::relationship::type() const```
Returns the type of this relationship.
#### ```xlnt::target_mode xlnt::relationship::target_mode() const```
Returns whether the target of the relationship is internal or external to the package.
#### ```uri xlnt::relationship::source() const```
Returns the URI of the package part this relationship points to.
#### ```uri xlnt::relationship::target() const```
Returns the URI of the package part this relationship points to.
#### ```bool xlnt::relationship::operator==(const relationship &rhs) const```
Returns true if and only if rhs is equal to this relationship.
#### ```bool xlnt::relationship::operator!=(const relationship &rhs) const```
Returns true if and only if rhs is not equal to this relationship.
### uri
#### ```xlnt::uri::uri()```
Constructs an empty URI.
#### ```xlnt::uri::uri(const uri &base, const uri &relative)```
Constructs a URI by combining base with relative.
#### ```xlnt::uri::uri(const uri &base, const path &relative)```
Constructs a URI by combining base with relative path.
#### ```xlnt::uri::uri(const std::string &uri_string)```
Constructs a URI by parsing the given uri_string.
#### ```bool xlnt::uri::is_relative() const```
Returns true if this URI is relative.
#### ```bool xlnt::uri::is_absolute() const```
Returns true if this URI is not relative (i.e. absolute).
#### ```std::string xlnt::uri::scheme() const```
Returns the scheme of this URI. E.g. the scheme of http://user:pass@example.com is "http"
#### ```std::string xlnt::uri::authority() const```
Returns the authority of this URI. E.g. the authority of http://user:pass@example.com:80/document is "user:pass@example.com:80"
#### ```bool xlnt::uri::has_authentication() const```
Returns true if an authentication section is specified for this URI.
#### ```std::string xlnt::uri::authentication() const```
Returns the authentication of this URI. E.g. the authentication of http://user:pass@example.com is "user:pass"
#### ```std::string xlnt::uri::username() const```
Returns the username of this URI. E.g. the username of http://user:pass@example.com is "user"
#### ```std::string xlnt::uri::password() const```
Returns the password of this URI. E.g. the password of http://user:pass@example.com is "pass"
#### ```std::string xlnt::uri::host() const```
Returns the host of this URI. E.g. the host of http://example.com:80/document is "example.com"
#### ```bool xlnt::uri::has_port() const```
Returns true if a non-default port is specified for this URI.
#### ```std::size_t xlnt::uri::port() const```
Returns the port of this URI. E.g. the port of https://example.com:443/document is "443"
#### ```class path xlnt::uri::path() const```
Returns the path of this URI. E.g. the path of http://example.com/document is "/document"
#### ```bool xlnt::uri::has_query() const```
Returns true if this URI has a non-null query string section.
#### ```std::string xlnt::uri::query() const```
Returns the query string of this URI. E.g. the query of http://example.com/document?v=1&x=3#abc is "v=1&x=3"
#### ```bool xlnt::uri::has_fragment() const```
Returns true if this URI has a non-empty fragment section.
#### ```std::string xlnt::uri::fragment() const```
Returns the fragment section of this URI. E.g. the fragment of http://example.com/document#abc is "abc"
#### ```std::string xlnt::uri::to_string() const```
Returns a string representation of this URI.
#### ```uri xlnt::uri::make_absolute(const uri &base)```
If this URI is relative, an absolute URI will be returned by appending the path to the given absolute base URI.
#### ```uri xlnt::uri::make_reference(const uri &base)```
If this URI is absolute, a relative URI will be returned by removing the common base path from the given absolute base URI.
#### ```bool xlnt::uri::operator==(const uri &other) const```
Returns true if this URI is equivalent to other.
## Styles Module
### alignment
#### ```bool xlnt::alignment::shrink() const```
Returns true if shrink-to-fit has been enabled.
#### ```alignment& xlnt::alignment::shrink(bool shrink_to_fit)```
Sets whether the font size should be reduced until all of the text fits in a cell without wrapping.
#### ```bool xlnt::alignment::wrap() const```
Returns true if text-wrapping has been enabled.
#### ```alignment& xlnt::alignment::wrap(bool wrap_text)```
Sets whether text in a cell should continue to multiple lines if it doesn't fit in one line.
#### ```optional<int> xlnt::alignment::indent() const```
Returns the optional value of indentation width in number of spaces.
#### ```alignment& xlnt::alignment::indent(int indent_size)```
Sets the indent size in number of spaces from the side of the cell. This will only take effect when left or right horizontal alignment has also been set.
#### ```optional<int> xlnt::alignment::rotation() const```
Returns the optional value of rotation for text in the cell in degrees.
#### ```alignment& xlnt::alignment::rotation(int text_rotation)```
Sets the rotation for text in the cell in degrees.
#### ```optional<horizontal_alignment> xlnt::alignment::horizontal() const```
Returns the optional horizontal alignment.
#### ```alignment& xlnt::alignment::horizontal(horizontal_alignment horizontal)```
Sets the horizontal alignment.
#### ```optional<vertical_alignment> xlnt::alignment::vertical() const```
Returns the optional vertical alignment.
#### ```alignment& xlnt::alignment::vertical(vertical_alignment vertical)```
Sets the vertical alignment.
#### ```bool xlnt::alignment::operator==(const alignment &other) const```
Returns true if this alignment is equivalent to other.
#### ```bool xlnt::alignment::operator!=(const alignment &other) const```
Returns true if this alignment is not equivalent to other.
### border
#### ```static const std::vector<border_side>& xlnt::border::all_sides()```
Returns a vector containing all of the border sides to be used for iteration.
#### ```xlnt::border::border()```
Constructs a default border.
#### ```optional<border_property> xlnt::border::side(border_side s) const```
Returns the border properties of the given side.
#### ```border& xlnt::border::side(border_side s, const border_property &prop)```
Sets the border properties of the side s to prop.
#### ```optional<diagonal_direction> xlnt::border::diagonal() const```
Returns the diagonal direction of this border.
#### ```border& xlnt::border::diagonal(diagonal_direction dir)```
Sets the diagonal direction of this border to dir.
#### ```bool xlnt::border::operator==(const border &right) const```
Returns true if left is exactly equal to right.
#### ```bool xlnt::border::operator!=(const border &right) const```
Returns true if left is not exactly equal to right.
### border_property
#### ```optional<class color> xlnt::border::border_property::color() const```
Returns the color of the side.
#### ```border_property& xlnt::border::border_property::color(const xlnt::color &c)```
Sets the color of the side and returns a reference to the side properties.
#### ```optional<border_style> xlnt::border::border_property::style() const```
Returns the style of the side.
#### ```border_property& xlnt::border::border_property::style(border_style style)```
Sets the style of the side and returns a reference to the side properties.
#### ```bool xlnt::border::border_property::operator==(const border_property &right) const```
Returns true if left is exactly equal to right.
#### ```bool xlnt::border::border_property::operator!=(const border_property &right) const```
Returns true if left is not exactly equal to right.
### indexed_color
#### ```xlnt::indexed_color::indexed_color(std::size_t index)```
Constructs an indexed_color from an index.
#### ```std::size_t xlnt::indexed_color::index() const```
Returns the index this color points to.
#### ```void xlnt::indexed_color::index(std::size_t index)```
Sets the index to index.
### theme_color
#### ```xlnt::theme_color::theme_color(std::size_t index)```
Constructs a theme_color from an index.
#### ```std::size_t xlnt::theme_color::index() const```
Returns the index of the color in the theme this points to.
#### ```void xlnt::theme_color::index(std::size_t index)```
Sets the index of this color to index.
### rgb_color
#### ```xlnt::rgb_color::rgb_color(const std::string &hex_string)```
Constructs an RGB color from a string in the form #[aa]rrggbb
#### ```xlnt::rgb_color::rgb_color(std::uint8_t r, std::uint8_t g, std::uint8_t b, std::uint8_t a=255)```
Constructs an RGB color from red, green, and blue values in the range 0 to 255 plus an optional alpha which defaults to fully opaque.
#### ```std::string xlnt::rgb_color::hex_string() const```
Returns a string representation of this color in the form #aarrggbb
#### ```std::uint8_t xlnt::rgb_color::red() const```
Returns a byte representing the red component of this color
#### ```std::uint8_t xlnt::rgb_color::green() const```
Returns a byte representing the red component of this color
#### ```std::uint8_t xlnt::rgb_color::blue() const```
Returns a byte representing the blue component of this color
#### ```std::uint8_t xlnt::rgb_color::alpha() const```
Returns a byte representing the alpha component of this color
#### ```std::array<std::uint8_t, 3> xlnt::rgb_color::rgb() const```
Returns the red, green, and blue components of this color separately in an array in that order.
#### ```std::array<std::uint8_t, 4> xlnt::rgb_color::rgba() const```
Returns the red, green, blue, and alpha components of this color separately in an array in that order.
### color
#### ```static const color xlnt::color::black()```
Returns the color #000000
#### ```static const color xlnt::color::white()```
Returns the color #ffffff
#### ```static const color xlnt::color::red()```
Returns the color #ff0000
#### ```static const color xlnt::color::darkred()```
Returns the color #8b0000
#### ```static const color xlnt::color::blue()```
Returns the color #00ff00
#### ```static const color xlnt::color::darkblue()```
Returns the color #008b00
#### ```static const color xlnt::color::green()```
Returns the color #0000ff
#### ```static const color xlnt::color::darkgreen()```
Returns the color #00008b
#### ```static const color xlnt::color::yellow()```
Returns the color #ffff00
#### ```static const color xlnt::color::darkyellow()```
Returns the color #cccc00
#### ```xlnt::color::color()```
Constructs a default color
#### ```xlnt::color::color(const rgb_color &rgb)```
Constructs a color from a given RGB color
#### ```xlnt::color::color(const indexed_color &indexed)```
Constructs a color from a given indexed color
#### ```xlnt::color::color(const theme_color &theme)```
Constructs a color from a given theme color
#### ```color_type xlnt::color::type() const```
Returns the type of this color
#### ```bool xlnt::color::auto_() const```
Returns true if this color has been set to auto
#### ```void xlnt::color::auto_(bool value)```
Sets the auto property of this color to value
#### ```rgb_color xlnt::color::rgb() const```
Returns the internal indexed color representing this color. If this is not an RGB color, an invalid_attribute exception will be thrown.
#### ```indexed_color xlnt::color::indexed() const```
Returns the internal indexed color representing this color. If this is not an indexed color, an invalid_attribute exception will be thrown.
#### ```theme_color xlnt::color::theme() const```
Returns the internal indexed color representing this color. If this is not a theme color, an invalid_attribute exception will be thrown.
#### ```double xlnt::color::tint() const```
Returns the tint of this color.
#### ```void xlnt::color::tint(double tint)```
Sets the tint of this color to tint. Tints lighten or darken an existing color by multiplying the color with the tint.
#### ```bool xlnt::color::operator==(const color &other) const```
Returns true if this color is equivalent to other
#### ```bool xlnt::color::operator!=(const color &other) const```
Returns true if this color is not equivalent to other
### pattern_fill
#### ```xlnt::pattern_fill::pattern_fill()```
Constructs a default pattern fill with a none pattern and no colors.
#### ```pattern_fill_type xlnt::pattern_fill::type() const```
Returns the pattern used by this fill
#### ```pattern_fill& xlnt::pattern_fill::type(pattern_fill_type new_type)```
Sets the pattern of this fill and returns a reference to it.
#### ```optional<color> xlnt::pattern_fill::foreground() const```
Returns the optional foreground color of this fill
#### ```pattern_fill& xlnt::pattern_fill::foreground(const color &foreground)```
Sets the foreground color and returns a reference to this pattern.
#### ```optional<color> xlnt::pattern_fill::background() const```
Returns the optional background color of this fill
#### ```pattern_fill& xlnt::pattern_fill::background(const color &background)```
Sets the foreground color and returns a reference to this pattern.
#### ```bool xlnt::pattern_fill::operator==(const pattern_fill &other) const```
Returns true if this pattern fill is equivalent to other.
#### ```bool xlnt::pattern_fill::operator!=(const pattern_fill &other) const```
Returns true if this pattern fill is not equivalent to other.
### gradient_fill
#### ```xlnt::gradient_fill::gradient_fill()```
Constructs a default linear fill
#### ```gradient_fill_type xlnt::gradient_fill::type() const```
Returns the type of this gradient fill
#### ```gradient_fill& xlnt::gradient_fill::type(gradient_fill_type new_type)```
Sets the type of this gradient fill
#### ```gradient_fill& xlnt::gradient_fill::degree(double degree)```
Sets the angle of the gradient in degrees
#### ```double xlnt::gradient_fill::degree() const```
Returns the angle of the gradient
#### ```double xlnt::gradient_fill::left() const```
Returns the distance from the left where the gradient starts.
#### ```gradient_fill& xlnt::gradient_fill::left(double value)```
Sets the distance from the left where the gradient starts.
#### ```double xlnt::gradient_fill::right() const```
Returns the distance from the right where the gradient starts.
#### ```gradient_fill& xlnt::gradient_fill::right(double value)```
Sets the distance from the right where the gradient starts.
#### ```double xlnt::gradient_fill::top() const```
Returns the distance from the top where the gradient starts.
#### ```gradient_fill& xlnt::gradient_fill::top(double value)```
Sets the distance from the top where the gradient starts.
#### ```double xlnt::gradient_fill::bottom() const```
Returns the distance from the bottom where the gradient starts.
#### ```gradient_fill& xlnt::gradient_fill::bottom(double value)```
Sets the distance from the bottom where the gradient starts.
#### ```gradient_fill& xlnt::gradient_fill::add_stop(double position, color stop_color)```
Adds a gradient stop at position with the given color.
#### ```gradient_fill& xlnt::gradient_fill::clear_stops()```
Deletes all stops from the gradient.
#### ```std::unordered_map<double, color> xlnt::gradient_fill::stops() const```
Returns all of the gradient stops.
#### ```bool xlnt::gradient_fill::operator==(const gradient_fill &other) const```
Returns true if the gradient is equivalent to other.
#### ```bool xlnt::gradient_fill::operator!=(const gradient_fill &right) const```
Returns true if the gradient is not equivalent to other.
### fill
#### ```static fill xlnt::fill::solid(const color &fill_color)```
Helper method for the most common use case, setting the fill color of a cell to a single solid color. The foreground and background colors of a fill are not the same as the foreground and background colors of a cell. When setting a fill color in Excel, a new fill is created with the given color as the fill's fgColor and index color number 64 as the bgColor. This method creates a fill in the same way.
#### ```xlnt::fill::fill()```
Constructs a fill initialized as a none-type pattern fill with no foreground or background colors.
#### ```xlnt::fill::fill(const pattern_fill &pattern)```
Constructs a fill initialized as a pattern fill based on the given pattern.
#### ```xlnt::fill::fill(const gradient_fill &gradient)```
Constructs a fill initialized as a gradient fill based on the given gradient.
#### ```fill_type xlnt::fill::type() const```
Returns the fill_type of this fill depending on how it was constructed.
#### ```class gradient_fill xlnt::fill::gradient_fill() const```
Returns the gradient fill represented by this fill. Throws an invalid_attribute exception if this is not a gradient fill.
#### ```class pattern_fill xlnt::fill::pattern_fill() const```
Returns the pattern fill represented by this fill. Throws an invalid_attribute exception if this is not a pattern fill.
#### ```bool xlnt::fill::operator==(const fill &other) const```
Returns true if left is exactly equal to right.
#### ```bool xlnt::fill::operator!=(const fill &other) const```
Returns true if left is not exactly equal to right.
### font
#### ```undefinedundefined```
Text can be underlined in the enumerated ways
#### ```xlnt::font::font()```
Constructs a default font. Calibri, size 12
#### ```font& xlnt::font::bold(bool bold)```
Sets the bold state of the font to bold and returns a reference to the font.
#### ```bool xlnt::font::bold() const```
Returns the bold state of the font.
#### ```font& xlnt::font::superscript(bool superscript)```
Sets the bold state of the font to bold and returns a reference to the font.
#### ```bool xlnt::font::superscript() const```
Returns true if this font has a superscript applied.
#### ```font& xlnt::font::italic(bool italic)```
Sets the bold state of the font to bold and returns a reference to the font.
#### ```bool xlnt::font::italic() const```
Returns true if this font has italics applied.
#### ```font& xlnt::font::strikethrough(bool strikethrough)```
Sets the bold state of the font to bold and returns a reference to the font.
#### ```bool xlnt::font::strikethrough() const```
Returns true if this font has a strikethrough applied.
#### ```font& xlnt::font::outline(bool outline)```
Sets the bold state of the font to bold and returns a reference to the font.
#### ```bool xlnt::font::outline() const```
Returns true if this font has an outline applied.
#### ```font& xlnt::font::shadow(bool shadow)```
Sets the shadow state of the font to shadow and returns a reference to the font.
#### ```bool xlnt::font::shadow() const```
Returns true if this font has a shadow applied.
#### ```font& xlnt::font::underline(underline_style new_underline)```
Sets the underline state of the font to new_underline and returns a reference to the font.
#### ```bool xlnt::font::underlined() const```
Returns true if this font has any type of underline applied.
#### ```underline_style xlnt::font::underline() const```
Returns the particular style of underline this font has applied.
#### ```bool xlnt::font::has_size() const```
Returns true if this font has a defined size.
#### ```font& xlnt::font::size(double size)```
Sets the size of the font to size and returns a reference to the font.
#### ```double xlnt::font::size() const```
Returns the size of the font.
#### ```bool xlnt::font::has_name() const```
Returns true if this font has a particular face applied (e.g. "Comic Sans").
#### ```font& xlnt::font::name(const std::string &name)```
Sets the font face to name and returns a reference to the font.
#### ```std::string xlnt::font::name() const```
Returns the name of the font face.
#### ```bool xlnt::font::has_color() const```
Returns true if this font has a color applied.
#### ```font& xlnt::font::color(const color &c)```
Sets the color of the font to c and returns a reference to the font.
#### ```xlnt::color xlnt::font::color() const```
Returns the color that this font is using.
#### ```bool xlnt::font::has_family() const```
Returns true if this font has a family defined.
#### ```font& xlnt::font::family(std::size_t family)```
Sets the family index of the font to family and returns a reference to the font.
#### ```std::size_t xlnt::font::family() const```
Returns the family index for the font.
#### ```bool xlnt::font::has_charset() const```
Returns true if this font has a charset defined.
#### ```font& xlnt::font::charset(std::size_t charset)```
Sets the charset of the font to charset and returns a reference to the font.
#### ```std::size_t xlnt::font::charset() const```
Returns the charset of the font.
#### ```bool xlnt::font::has_scheme() const```
Returns true if this font has a scheme.
#### ```font& xlnt::font::scheme(const std::string &scheme)```
Sets the scheme of the font to scheme and returns a reference to the font.
#### ```std::string xlnt::font::scheme() const```
Returns the scheme of this font.
#### ```bool xlnt::font::operator==(const font &other) const```
Returns true if left is exactly equal to right.
#### ```bool xlnt::font::operator!=(const font &other) const```
Returns true if left is not exactly equal to right.
### format
#### ```friend struct detail::stylesheetundefined```
#### ```friend class detail::xlsx_producerundefined```
#### ```class alignment xlnt::format::alignment() const```
Returns the alignment of this format.
#### ```format xlnt::format::alignment(const xlnt::alignment &new_alignment, bool applied=true)```
Sets the alignment of this format to new_alignment. Applied, which defaults to true, determines whether the alignment should be enabled for cells using this format.
#### ```bool xlnt::format::alignment_applied() const```
Returns true if the alignment of this format should be applied to cells using it.
#### ```class border xlnt::format::border() const```
Returns the border of this format.
#### ```format xlnt::format::border(const xlnt::border &new_border, bool applied=true)```
Sets the border of this format to new_border. Applied, which defaults to true, determines whether the border should be enabled for cells using this format.
#### ```bool xlnt::format::border_applied() const```
Returns true if the border set for this format should be applied to cells using the format.
#### ```class fill xlnt::format::fill() const```
Returns the fill of this format.
#### ```format xlnt::format::fill(const xlnt::fill &new_fill, bool applied=true)```
Sets the fill of this format to new_fill. Applied, which defaults to true, determines whether the border should be enabled for cells using this format.
#### ```bool xlnt::format::fill_applied() const```
Returns true if the fill set for this format should be applied to cells using the format.
#### ```class font xlnt::format::font() const```
Returns the font of this format.
#### ```format xlnt::format::font(const xlnt::font &new_font, bool applied=true)```
Sets the font of this format to new_font. Applied, which defaults to true, determines whether the font should be enabled for cells using this format.
#### ```bool xlnt::format::font_applied() const```
Returns true if the font set for this format should be applied to cells using the format.
#### ```class number_format xlnt::format::number_format() const```
Returns the number format of this format.
#### ```format xlnt::format::number_format(const xlnt::number_format &new_number_format, bool applied=true)```
Sets the number format of this format to new_number_format. Applied, which defaults to true, determines whether the number format should be enabled for cells using this format.
#### ```bool xlnt::format::number_format_applied() const```
Returns true if the number_format set for this format should be applied to cells using the format.
#### ```class protection xlnt::format::protection() const```
Returns the protection of this format.
#### ```bool xlnt::format::protection_applied() const```
Returns true if the protection set for this format should be applied to cells using the format.
#### ```format xlnt::format::protection(const xlnt::protection &new_protection, bool applied=true)```
Sets the protection of this format to new_protection. Applied, which defaults to true, determines whether the protection should be enabled for cells using this format.
#### ```bool xlnt::format::pivot_button() const```
Returns true if the pivot table interface is enabled for this format.
#### ```void xlnt::format::pivot_button(bool show)```
If show is true, a pivot table interface will be displayed for cells using this format.
#### ```bool xlnt::format::quote_prefix() const```
Returns true if this format should add a single-quote prefix for all text values.
#### ```void xlnt::format::quote_prefix(bool quote)```
If quote is true, enables a single-quote prefix for all text values in cells using this format (e.g. "abc" will appear as "'abc"). The text will also not be stored in sharedStrings when this is enabled.
#### ```bool xlnt::format::has_style() const```
Returns true if this format has a corresponding style applied.
#### ```void xlnt::format::clear_style()```
Removes the style from this format if it exists.
#### ```format xlnt::format::style(const std::string &name)```
Sets the style of this format to a style with the given name.
#### ```format xlnt::format::style(const class style &new_style)```
Sets the style of this format to new_style.
#### ```class style xlnt::format::style()```
Returns the style of this format. If it has no style, an invalid_parameter exception will be thrown.
#### ```const class style xlnt::format::style() const```
Returns the style of this format. If it has no style, an invalid_parameters exception will be thrown.
### number_format
#### ```static const number_format xlnt::number_format::general()```
Number format "General"
#### ```static const number_format xlnt::number_format::text()```
Number format "@"
#### ```static const number_format xlnt::number_format::number()```
Number format "0"
#### ```static const number_format xlnt::number_format::number_00()```
Number format "00"
#### ```static const number_format xlnt::number_format::number_comma_separated1()```
Number format "#,##0.00"
#### ```static const number_format xlnt::number_format::percentage()```
Number format "0%"
#### ```static const number_format xlnt::number_format::percentage_00()```
Number format "0.00%"
#### ```static const number_format xlnt::number_format::date_yyyymmdd2()```
Number format "yyyy-mm-dd"
#### ```static const number_format xlnt::number_format::date_yymmdd()```
Number format "yy-mm-dd"
#### ```static const number_format xlnt::number_format::date_ddmmyyyy()```
Number format "dd/mm/yy"
#### ```static const number_format xlnt::number_format::date_dmyslash()```
Number format "d/m/yy"
#### ```static const number_format xlnt::number_format::date_dmyminus()```
Number format "d-m-yy"
#### ```static const number_format xlnt::number_format::date_dmminus()```
Number format "d-m"
#### ```static const number_format xlnt::number_format::date_myminus()```
Number format "m-yy"
#### ```static const number_format xlnt::number_format::date_xlsx14()```
Number format "mm-dd-yy"
#### ```static const number_format xlnt::number_format::date_xlsx15()```
Number format "d-mmm-yy"
#### ```static const number_format xlnt::number_format::date_xlsx16()```
Number format "d-mmm"
#### ```static const number_format xlnt::number_format::date_xlsx17()```
Number format "mmm-yy"
#### ```static const number_format xlnt::number_format::date_xlsx22()```
Number format "m/d/yy h:mm"
#### ```static const number_format xlnt::number_format::date_datetime()```
Number format "yyyy-mm-dd h:mm:ss"
#### ```static const number_format xlnt::number_format::date_time1()```
Number format "h:mm AM/PM"
#### ```static const number_format xlnt::number_format::date_time2()```
Number format "h:mm:ss AM/PM"
#### ```static const number_format xlnt::number_format::date_time3()```
Number format "h:mm"
#### ```static const number_format xlnt::number_format::date_time4()```
Number format "h:mm:ss"
#### ```static const number_format xlnt::number_format::date_time5()```
Number format "mm:ss"
#### ```static const number_format xlnt::number_format::date_time6()```
Number format "h:mm:ss"
#### ```static bool xlnt::number_format::is_builtin_format(std::size_t builtin_id)```
Returns true if the given format ID corresponds to a known builtin format.
#### ```static const number_format& xlnt::number_format::from_builtin_id(std::size_t builtin_id)```
Returns the format with the given ID. Thows an invalid_parameter exception if builtin_id is not a valid ID.
#### ```xlnt::number_format::number_format()```
Constructs a default number_format equivalent to "General"
#### ```xlnt::number_format::number_format(std::size_t builtin_id)```
Constructs a number format equivalent to that returned from number_format::from_builtin_id(builtin_id).
#### ```xlnt::number_format::number_format(const std::string &code)```
Constructs a number format from a code string. If the string matches a builtin ID, its ID will also be set to match the builtin ID.
#### ```xlnt::number_format::number_format(const std::string &code, std::size_t custom_id)```
Constructs a number format from a code string and custom ID. Custom ID should generally be >= 164.
#### ```void xlnt::number_format::format_string(const std::string &format_code)```
Sets the format code of this number format to format_code.
#### ```void xlnt::number_format::format_string(const std::string &format_code, std::size_t custom_id)```
Sets the format code of this number format to format_code and the ID to custom_id.
#### ```std::string xlnt::number_format::format_string() const```
Returns the format code this number format uses.
#### ```bool xlnt::number_format::has_id() const```
Returns true if this number format has an ID.
#### ```void xlnt::number_format::id(std::size_t id)```
Sets the ID of this number format to id.
#### ```std::size_t xlnt::number_format::id() const```
Returns the ID of this format.
#### ```std::string xlnt::number_format::format(const std::string &text) const```
Returns text formatted according to this number format's format code.
#### ```std::string xlnt::number_format::format(long double number, calendar base_date) const```
Returns number formatted according to this number format's format code with the given base date.
#### ```bool xlnt::number_format::is_date_format() const```
Returns true if this format code returns a number formatted as a date.
#### ```bool xlnt::number_format::operator==(const number_format &other) const```
Returns true if this format is equivalent to other.
#### ```bool xlnt::number_format::operator!=(const number_format &other) const```
Returns true if this format is not equivalent to other.
### protection
#### ```static protection xlnt::protection::unlocked_and_visible()```
Returns an unlocked and unhidden protection object.
#### ```static protection xlnt::protection::locked_and_visible()```
Returns a locked and unhidden protection object.
#### ```static protection xlnt::protection::unlocked_and_hidden()```
Returns an unlocked and hidden protection object.
#### ```static protection xlnt::protection::locked_and_hidden()```
Returns a locked and hidden protection object.
#### ```xlnt::protection::protection()```
Constructs a default unlocked unhidden protection object.
#### ```bool xlnt::protection::locked() const```
Returns true if cells using this protection should be locked.
#### ```protection& xlnt::protection::locked(bool locked)```
Sets the locked state of the protection to locked and returns a reference to the protection.
#### ```bool xlnt::protection::hidden() const```
Returns true if cells using this protection should be hidden.
#### ```protection& xlnt::protection::hidden(bool hidden)```
Sets the hidden state of the protection to hidden and returns a reference to the protection.
#### ```bool xlnt::protection::operator==(const protection &other) const```
Returns true if this protections is equivalent to right.
#### ```bool xlnt::protection::operator!=(const protection &other) const```
Returns true if this protection is not equivalent to right.
### style
#### ```friend struct detail::stylesheetundefined```
#### ```friend class detail::xlsx_consumerundefined```
#### ```xlnt::style::style()=delete```
Delete zero-argument constructor
#### ```xlnt::style::style(const style &other)=default```
Default copy constructor. Constructs a style using the same PIMPL as other.
#### ```std::string xlnt::style::name() const```
Returns the name of this style.
#### ```style xlnt::style::name(const std::string &name)```
Sets the name of this style to name.
#### ```bool xlnt::style::hidden() const```
Returns true if this style is hidden.
#### ```style xlnt::style::hidden(bool value)```
Sets the hidden state of this style to value. A hidden style will not be shown in the list of selectable styles in the UI, but will still apply its formatting to cells using it.
#### ```bool xlnt::style::custom_builtin() const```
Returns true if this is a builtin style that has been customized and should therefore be persisted in the workbook.
#### ```std::size_t xlnt::style::builtin_id() const```
Returns the index of the builtin style that this style is an instance of or is a customized version thereof. If style::builtin() is false, this will throw an invalid_attribute exception.
#### ```bool xlnt::style::builtin() const```
Returns true if this is a builtin style.
#### ```class alignment xlnt::style::alignment() const```
Returns the alignment of this style.
#### ```bool xlnt::style::alignment_applied() const```
Returns true if the alignment of this style should be applied to cells using it.
#### ```style xlnt::style::alignment(const xlnt::alignment &new_alignment, bool applied=true)```
Sets the alignment of this style to new_alignment. Applied, which defaults to true, determines whether the alignment should be enabled for cells using this style.
#### ```class border xlnt::style::border() const```
Returns the border of this style.
#### ```bool xlnt::style::border_applied() const```
Returns true if the border set for this style should be applied to cells using the style.
#### ```style xlnt::style::border(const xlnt::border &new_border, bool applied=true)```
Sets the border of this style to new_border. Applied, which defaults to true, determines whether the border should be enabled for cells using this style.
#### ```class fill xlnt::style::fill() const```
Returns the fill of this style.
#### ```bool xlnt::style::fill_applied() const```
Returns true if the fill set for this style should be applied to cells using the style.
#### ```style xlnt::style::fill(const xlnt::fill &new_fill, bool applied=true)```
Sets the fill of this style to new_fill. Applied, which defaults to true, determines whether the border should be enabled for cells using this style.
#### ```class font xlnt::style::font() const```
Returns the font of this style.
#### ```bool xlnt::style::font_applied() const```
Returns true if the font set for this style should be applied to cells using the style.
#### ```style xlnt::style::font(const xlnt::font &new_font, bool applied=true)```
Sets the font of this style to new_font. Applied, which defaults to true, determines whether the font should be enabled for cells using this style.
#### ```class number_format xlnt::style::number_format() const```
Returns the number_format of this style.
#### ```bool xlnt::style::number_format_applied() const```
Returns true if the number_format set for this style should be applied to cells using the style.
#### ```style xlnt::style::number_format(const xlnt::number_format &new_number_format, bool applied=true)```
Sets the number format of this style to new_number_format. Applied, which defaults to true, determines whether the number format should be enabled for cells using this style.
#### ```class protection xlnt::style::protection() const```
Returns the protection of this style.
#### ```bool xlnt::style::protection_applied() const```
Returns true if the protection set for this style should be applied to cells using the style.
#### ```style xlnt::style::protection(const xlnt::protection &new_protection, bool applied=true)```
Sets the border of this style to new_protection. Applied, which defaults to true, determines whether the protection should be enabled for cells using this style.
#### ```bool xlnt::style::pivot_button() const```
Returns true if the pivot table interface is enabled for this style.
#### ```void xlnt::style::pivot_button(bool show)```
If show is true, a pivot table interface will be displayed for cells using this style.
#### ```bool xlnt::style::quote_prefix() const```
Returns true if this style should add a single-quote prefix for all text values.
#### ```void xlnt::style::quote_prefix(bool quote)```
If quote is true, enables a single-quote prefix for all text values in cells using this style (e.g. "abc" will appear as "'abc"). The text will also not be stored in sharedStrings when this is enabled.
#### ```bool xlnt::style::operator==(const style &other) const```
Returns true if this style is equivalent to other.
#### ```bool xlnt::style::operator!=(const style &other) const```
Returns true if this style is not equivalent to other.
## Utils Module
### date
#### ```int xlnt::date::yearundefined```
The year
#### ```int xlnt::date::monthundefined```
The month
#### ```int xlnt::date::dayundefined```
The day
#### ```static date xlnt::date::today()```
Return the current date according to the system time.
#### ```static date xlnt::date::from_number(int days_since_base_year, calendar base_date)```
Return a date by adding days_since_base_year to base_date. This includes leap years.
#### ```xlnt::date::date(int year_, int month_, int day_)```
Constructs a data from a given year, month, and day.
#### ```int xlnt::date::to_number(calendar base_date) const```
Return the number of days between this date and base_date.
#### ```int xlnt::date::weekday() const```
Calculates and returns the day of the week that this date represents in the range 0 to 6 where 0 represents Sunday.
#### ```bool xlnt::date::operator==(const date &comparand) const```
Return true if this date is equal to comparand.
### datetime
#### ```int xlnt::datetime::yearundefined```
The year
#### ```int xlnt::datetime::monthundefined```
The month
#### ```int xlnt::datetime::dayundefined```
The day
#### ```int xlnt::datetime::hourundefined```
The hour
#### ```int xlnt::datetime::minuteundefined```
The minute
#### ```int xlnt::datetime::secondundefined```
The second
#### ```int xlnt::datetime::microsecondundefined```
The microsecond
#### ```static datetime xlnt::datetime::now()```
Returns the current date and time according to the system time.
#### ```static datetime xlnt::datetime::today()```
Returns the current date and time according to the system time. This is equivalent to datetime::now().
#### ```static datetime xlnt::datetime::from_number(long double number, calendar base_date)```
Returns a datetime from number by converting the integer part into a date and the fractional part into a time according to date::from_number and time::from_number.
#### ```static datetime xlnt::datetime::from_iso_string(const std::string &iso_string)```
Returns a datetime equivalent to the ISO-formatted string iso_string.
#### ```xlnt::datetime::datetime(const date &d, const time &t)```
Constructs a datetime from a date and a time.
#### ```xlnt::datetime::datetime(int year_, int month_, int day_, int hour_=0, int minute_=0, int second_=0, int microsecond_=0)```
Constructs a datetime from a year, month, and day plus optional hour, minute, second, and microsecond.
#### ```std::string xlnt::datetime::to_string() const```
Returns a string represenation of this date and time.
#### ```std::string xlnt::datetime::to_iso_string() const```
Returns an ISO-formatted string representation of this date and time.
#### ```long double xlnt::datetime::to_number(calendar base_date) const```
Returns this datetime as a number of seconds since 1900 or 1904 (depending on base_date provided).
#### ```bool xlnt::datetime::operator==(const datetime &comparand) const```
Returns true if this datetime is equivalent to comparand.
#### ```int xlnt::datetime::weekday() const```
Calculates and returns the day of the week that this date represents in the range 0 to 6 where 0 represents Sunday.
### exception
#### ```xlnt::exception::exception(const std::string &message)```
Constructs an exception with a message. This message will be returned by std::exception::what(), an inherited member of this class.
#### ```xlnt::exception::exception(const exception &)=default```
Default copy constructor.
#### ```virtual xlnt::exception::~exception()```
Destructor
#### ```void xlnt::exception::message(const std::string &message)```
Sets the message after the xlnt::exception is constructed. This can show more specific information than std::exception::what().
### invalid_parameter
#### ```xlnt::invalid_parameter::invalid_parameter()```
Default constructor.
#### ```xlnt::invalid_parameter::invalid_parameter(const invalid_parameter &)=default```
Default copy constructor.
#### ```virtual xlnt::invalid_parameter::~invalid_parameter()```
Destructor
### invalid_sheet_title
#### ```xlnt::invalid_sheet_title::invalid_sheet_title(const std::string &title)```
Default constructor.
#### ```xlnt::invalid_sheet_title::invalid_sheet_title(const invalid_sheet_title &)=default```
Default copy constructor.
#### ```virtual xlnt::invalid_sheet_title::~invalid_sheet_title()```
Destructor
### missing_number_format
#### ```xlnt::missing_number_format::missing_number_format()```
Default constructor.
#### ```xlnt::missing_number_format::missing_number_format(const missing_number_format &)=default```
Default copy constructor.
#### ```virtual xlnt::missing_number_format::~missing_number_format()```
Destructor
### invalid_file
#### ```xlnt::invalid_file::invalid_file(const std::string &filename)```
Constructs an invalid_file exception thrown when attempt to access the given filename.
#### ```xlnt::invalid_file::invalid_file(const invalid_file &)=default```
Default copy constructor.
#### ```virtual xlnt::invalid_file::~invalid_file()```
Destructor
### illegal_character
#### ```xlnt::illegal_character::illegal_character(char c)```
Constructs an illegal_character exception thrown as a result of the given character.
#### ```xlnt::illegal_character::illegal_character(const illegal_character &)=default```
Default copy constructor.
#### ```virtual xlnt::illegal_character::~illegal_character()```
Destructor
### invalid_data_type
#### ```xlnt::invalid_data_type::invalid_data_type()```
Default constructor.
#### ```xlnt::invalid_data_type::invalid_data_type(const invalid_data_type &)=default```
Default copy constructor.
#### ```virtual xlnt::invalid_data_type::~invalid_data_type()```
Destructor
### invalid_column_index
#### ```xlnt::invalid_column_index::invalid_column_index()```
Default constructor.
#### ```xlnt::invalid_column_index::invalid_column_index(const invalid_column_index &)=default```
Default copy constructor.
#### ```virtual xlnt::invalid_column_index::~invalid_column_index()```
Destructor
### invalid_cell_reference
#### ```xlnt::invalid_cell_reference::invalid_cell_reference(column_t column, row_t row)```
Constructs an invalid_cell_reference exception for the given column and row.
#### ```xlnt::invalid_cell_reference::invalid_cell_reference(const std::string &reference_string)```
Constructs an invalid_cell_reference exception for the given string.
#### ```xlnt::invalid_cell_reference::invalid_cell_reference(const invalid_cell_reference &)=default```
Default copy constructor.
#### ```virtual xlnt::invalid_cell_reference::~invalid_cell_reference()```
Destructor
### invalid_attribute
#### ```xlnt::invalid_attribute::invalid_attribute()```
Default constructor.
#### ```xlnt::invalid_attribute::invalid_attribute(const invalid_attribute &)=default```
Default copy constructor.
#### ```virtual xlnt::invalid_attribute::~invalid_attribute()```
Destructor
### key_not_found
#### ```xlnt::key_not_found::key_not_found()```
Default constructor.
#### ```xlnt::key_not_found::key_not_found(const key_not_found &)=default```
Default copy constructor.
#### ```virtual xlnt::key_not_found::~key_not_found()```
Destructor
### no_visible_worksheets
#### ```xlnt::no_visible_worksheets::no_visible_worksheets()```
Default constructor.
#### ```xlnt::no_visible_worksheets::no_visible_worksheets(const no_visible_worksheets &)=default```
Default copy constructor.
#### ```virtual xlnt::no_visible_worksheets::~no_visible_worksheets()```
Destructor
### unhandled_switch_case
#### ```xlnt::unhandled_switch_case::unhandled_switch_case()```
Default constructor.
#### ```xlnt::unhandled_switch_case::unhandled_switch_case(const unhandled_switch_case &)=default```
Default copy constructor.
#### ```virtual xlnt::unhandled_switch_case::~unhandled_switch_case()```
Destructor
### unsupported
#### ```xlnt::unsupported::unsupported(const std::string &message)```
Constructs an unsupported exception with a message describing the unsupported feature.
#### ```xlnt::unsupported::unsupported(const unsupported &)=default```
Default copy constructor.
#### ```virtual xlnt::unsupported::~unsupported()```
Destructor
### optional
#### ```xlnt::optional< T >::optional()```
Default contructor. is_set() will be false initially.
#### ```xlnt::optional< T >::optional(const T &value)```
Constructs this optional with a value.
#### ```bool xlnt::optional< T >::is_set() const```
Returns true if this object currently has a value set. This should be called before accessing the value with optional::get().
#### ```void xlnt::optional< T >::set(const T &value)```
Sets the value to value.
#### ```T& xlnt::optional< T >::get()```
Gets the value. If no value has been initialized in this object, an xlnt::invalid_attribute exception will be thrown.
#### ```const T& xlnt::optional< T >::get() const```
Gets the value. If no value has been initialized in this object, an xlnt::invalid_attribute exception will be thrown.
#### ```void xlnt::optional< T >::clear()```
Resets the internal value using its default constructor. After this is called, is_set() will return false until a new value is provided.
#### ```optional& xlnt::optional< T >::operator=(const T &rhs)```
Assignment operator. Equivalent to setting the value using optional::set.
#### ```bool xlnt::optional< T >::operator==(const optional< T > &other) const```
Returns true if neither this nor other have a value or both have a value and those values are equal according to their equality operator.
### path
#### ```static char xlnt::path::system_separator()```
The system-specific path separator character (e.g. '/' or '\').
#### ```xlnt::path::path()```
Construct an empty path.
#### ```xlnt::path::path(const std::string &path_string)```
Counstruct a path from a string representing the path.
#### ```xlnt::path::path(const std::string &path_string, char sep)```
Construct a path from a string with an explicit directory seprator.
#### ```bool xlnt::path::is_relative() const```
Return true iff this path doesn't begin with / (or a drive letter on Windows).
#### ```bool xlnt::path::is_absolute() const```
Return true iff path::is_relative() is false.
#### ```bool xlnt::path::is_root() const```
Return true iff this path is the root directory.
#### ```path xlnt::path::parent() const```
Return a new path that points to the directory containing the current path Return the path unchanged if this path is the absolute or relative root.
#### ```std::string xlnt::path::filename() const```
Return the last component of this path.
#### ```std::string xlnt::path::extension() const```
Return the part of the path following the last dot in the filename.
#### ```std::pair<std::string, std::string> xlnt::path::split_extension() const```
Return a pair of strings resulting from splitting the filename on the last dot.
#### ```std::vector<std::string> xlnt::path::split() const```
Create a string representing this path separated by the provided separator or the system-default separator if not provided.
#### ```std::string xlnt::path::string() const```
Create a string representing this path separated by the provided separator or the system-default separator if not provided.
#### ```path xlnt::path::resolve(const path &base_path) const```
If this path is relative, append each component of this path to base_path and return the resulting absolute path. Otherwise, the the current path will be returned and base_path will be ignored.
#### ```path xlnt::path::relative_to(const path &base_path) const```
The inverse of path::resolve. Creates a relative path from an absolute path by removing the common root between base_path and this path. If the current path is already relative, return it unchanged.
#### ```bool xlnt::path::exists() const```
Return true iff the file or directory pointed to by this path exists on the filesystem.
#### ```bool xlnt::path::is_directory() const```
Return true if the file or directory pointed to by this path is a directory.
#### ```bool xlnt::path::is_file() const```
Return true if the file or directory pointed to by this path is a regular file.
#### ```std::string xlnt::path::read_contents() const```
Open the file pointed to by this path and return a string containing the files contents.
#### ```path xlnt::path::append(const std::string &to_append) const```
Append the provided part to this path and return the result.
#### ```path xlnt::path::append(const path &to_append) const```
Append the provided part to this path and return the result.
#### ```bool xlnt::path::operator==(const path &other) const```
Returns true if left path is equal to right path.
### path >
#### ```size_t std::hash< xlnt::path >::operator()(const xlnt::path &p) const```
Returns a hashed represenation of the given path.
### scoped_enum_hash
#### ```std::size_t xlnt::scoped_enum_hash< Enum >::operator()(Enum e) const```
Cast the enumeration e to a std::size_t and hash that value using std::hash.
### time
#### ```int xlnt::time::hourundefined```
The hour
#### ```int xlnt::time::minuteundefined```
The minute
#### ```int xlnt::time::secondundefined```
The second
#### ```int xlnt::time::microsecondundefined```
The microsecond
#### ```static time xlnt::time::now()```
Return the current time according to the system time.
#### ```static time xlnt::time::from_number(long double number)```
Return a time from a number representing a fraction of a day. The integer part of number will be ignored. 0.5 would return time(12, 0, 0, 0) or noon, halfway through the day.
#### ```xlnt::time::time(int hour_=0, int minute_=0, int second_=0, int microsecond_=0)```
Constructs a time object from an optional hour, minute, second, and microsecond.
#### ```xlnt::time::time(const std::string &time_string)```
Constructs a time object from a string representing the time.
#### ```long double xlnt::time::to_number() const```
Returns a numeric representation of the time in the range 0-1 where the value is equal to the fraction of the day elapsed.
#### ```bool xlnt::time::operator==(const time &comparand) const```
Returns true if this time is equivalent to comparand.
### timedelta
#### ```int xlnt::timedelta::daysundefined```
The days
#### ```int xlnt::timedelta::hoursundefined```
The hours
#### ```int xlnt::timedelta::minutesundefined```
The minutes
#### ```int xlnt::timedelta::secondsundefined```
The seconds
#### ```int xlnt::timedelta::microsecondsundefined```
The microseconds
#### ```static timedelta xlnt::timedelta::from_number(long double number)```
Returns a timedelta from a number representing the factional number of days elapsed.
#### ```xlnt::timedelta::timedelta()```
Constructs a timedelta equal to zero.
#### ```xlnt::timedelta::timedelta(int days_, int hours_, int minutes_, int seconds_, int microseconds_)```
Constructs a timedelta from a number of days, hours, minutes, seconds, and microseconds.
#### ```long double xlnt::timedelta::to_number() const```
Returns a numeric representation of this timedelta as a fractional number of days.
### variant
#### ```undefinedundefined```
The possible types a variant can hold.
#### ```xlnt::variant::variant()```
Default constructor. Creates a null-type variant.
#### ```xlnt::variant::variant(const std::string &value)```
Creates a string-type variant with the given value.
#### ```xlnt::variant::variant(const char *value)```
Creates a string-type variant with the given value.
#### ```xlnt::variant::variant(int value)```
Creates a i4-type variant with the given value.
#### ```xlnt::variant::variant(bool value)```
Creates a bool-type variant with the given value.
#### ```xlnt::variant::variant(const datetime &value)```
Creates a date-type variant with the given value.
#### ```xlnt::variant::variant(const std::initializer_list< int > &value)```
Creates a vector_i4-type variant with the given value.
#### ```xlnt::variant::variant(const std::vector< int > &value)```
Creates a vector_i4-type variant with the given value.
#### ```xlnt::variant::variant(const std::initializer_list< const char *> &value)```
Creates a vector_string-type variant with the given value.
#### ```xlnt::variant::variant(const std::vector< const char *> &value)```
Creates a vector_string-type variant with the given value.
#### ```xlnt::variant::variant(const std::initializer_list< std::string > &value)```
Creates a vector_string-type variant with the given value.
#### ```xlnt::variant::variant(const std::vector< std::string > &value)```
Creates a vector_string-type variant with the given value.
#### ```xlnt::variant::variant(const std::vector< variant > &value)```
Creates a vector_variant-type variant with the given value.
#### ```bool xlnt::variant::is(type t) const```
Returns true if this variant is of type t.
#### ```T xlnt::variant::get() const```
Returns the value of this variant as type T. An exception will be thrown if the types are not convertible.
#### ```type xlnt::variant::value_type() const```
Returns the type of this variant.
## Workbook Module
### calculation_properties
#### ```std::size_t xlnt::calculation_properties::calc_idundefined```
The version of calculation engine used to calculate cell formula values. If this is older than the version of the Excel calculation engine opening the workbook, cell values will be recalculated.
#### ```bool xlnt::calculation_properties::concurrent_calcundefined```
If this is true, concurrent calculation will be enabled for the workbook.
### document_security
#### ```bool xlnt::document_security::lock_revisionundefined```
If true, the workbook is locked for revisions.
#### ```bool xlnt::document_security::lock_structureundefined```
If true, worksheets can't be moved, renamed, (un)hidden, inserted, or deleted.
#### ```bool xlnt::document_security::lock_windowsundefined```
If true, workbook windows will be opened at the same position with the same size every time they are loaded.
#### ```lock_verifier xlnt::document_security::revision_lockundefined```
The settings to allow the revision lock to be removed.
#### ```lock_verifier xlnt::document_security::workbook_lockundefined```
The settings to allow the structure and windows lock to be removed.
#### ```xlnt::document_security::document_security()```
Constructs a new document security object with default values.
### lock_verifier
#### ```std::string xlnt::document_security::lock_verifier::hash_algorithmundefined```
The algorithm used to create and verify this lock.
#### ```std::string xlnt::document_security::lock_verifier::saltundefined```
The initial salt combined with the password used to prevent rainbow table attacks
#### ```std::string xlnt::document_security::lock_verifier::hashundefined```
The actual hash value represented as a string
#### ```std::size_t xlnt::document_security::lock_verifier::spin_countundefined```
The number of times the hash should be applied to the password combined with the salt. This allows the difficulty of the hash to be increased as computing power increases.
### external_book
### named_range
#### ```using xlnt::named_range::target =  std::pair<worksheet, range_reference>undefined```
Type alias for the combination of sheet and range this named_range points to.
#### ```xlnt::named_range::named_range()```
Constructs a named range that has no name and has no targets.
#### ```xlnt::named_range::named_range(const named_range &other)```
Constructs a named range by copying its name and targets from other.
#### ```xlnt::named_range::named_range(const std::string &name, const std::vector< target > &targets)```
Constructs a named range with the given name and targets.
#### ```std::string xlnt::named_range::name() const```
Returns the name of this range.
#### ```const std::vector<target>& xlnt::named_range::targets() const```
Returns the set of targets of this named range as a vector.
#### ```named_range& xlnt::named_range::operator=(const named_range &other)```
Assigns the name and targets of this named_range to that of other.
### theme
### workbook
#### ```using xlnt::workbook::iterator =  worksheet_iteratorundefined```
typedef for the iterator used for iterating through this workbook (non-const) in a range-based for loop.
#### ```using xlnt::workbook::const_iterator =  const_worksheet_iteratorundefined```
typedef for the iterator used for iterating through this workbook (const) in a range-based for loop.
#### ```using xlnt::workbook::reverse_iterator =  std::reverse_iterator<iterator>undefined```
typedef for the iterator used for iterating through this workbook (non-const) in a range-based for loop in reverse order using std::make_reverse_iterator.
#### ```using xlnt::workbook::const_reverse_iterator =  std::reverse_iterator<const_iterator>undefined```
typedef for the iterator used for iterating through this workbook (const) in a range-based for loop in reverse order using std::make_reverse_iterator.
#### ```friend class detail::xlsx_consumerundefined```
#### ```friend class detail::xlsx_producerundefined```
#### ```static workbook xlnt::workbook::empty()```
Constructs and returns an empty workbook similar to a default. Excel workbook
#### ```xlnt::workbook::workbook()```
Default constructor. Constructs a workbook containing a single empty worksheet using workbook::empty().
#### ```xlnt::workbook::workbook(workbook &&other)```
Move constructor. Constructs a workbook from existing workbook, other.
#### ```xlnt::workbook::workbook(const workbook &other)```
Copy constructor. Constructs this workbook from existing workbook, other.
#### ```xlnt::workbook::~workbook()```
Destroys this workbook, deallocating all internal storage space. Any pimpl wrapper classes (e.g. cell) pointing into this workbook will be invalid after this is executed.
#### ```worksheet xlnt::workbook::create_sheet()```
Creates and returns a sheet after the last sheet in this workbook.
#### ```worksheet xlnt::workbook::create_sheet(std::size_t index)```
Creates and returns a sheet at the specified index.
#### ```worksheet xlnt::workbook::create_sheet_with_rel(const std::string &title, const relationship &rel)```
TODO: This should be private...
#### ```worksheet xlnt::workbook::copy_sheet(worksheet worksheet)```
Creates and returns a new sheet after the last sheet initializing it with all of the data from the provided worksheet.
#### ```worksheet xlnt::workbook::copy_sheet(worksheet worksheet, std::size_t index)```
Creates and returns a new sheet at the specified index initializing it with all of the data from the provided worksheet.
#### ```worksheet xlnt::workbook::active_sheet()```
Returns the worksheet that is determined to be active. An active sheet is that which is initially shown by the spreadsheet editor.
#### ```worksheet xlnt::workbook::sheet_by_title(const std::string &title)```
Returns the worksheet with the given name. This may throw an exception if the sheet isn't found. Use workbook::contains(const std::string &) to make sure the sheet exists before calling this method.
#### ```const worksheet xlnt::workbook::sheet_by_title(const std::string &title) const```
Returns the worksheet with the given name. This may throw an exception if the sheet isn't found. Use workbook::contains(const std::string &) to make sure the sheet exists before calling this method.
#### ```worksheet xlnt::workbook::sheet_by_index(std::size_t index)```
Returns the worksheet at the given index. This will throw an exception if index is greater than or equal to the number of sheets in this workbook.
#### ```const worksheet xlnt::workbook::sheet_by_index(std::size_t index) const```
Returns the worksheet at the given index. This will throw an exception if index is greater than or equal to the number of sheets in this workbook.
#### ```worksheet xlnt::workbook::sheet_by_id(std::size_t id)```
Returns the worksheet with a sheetId of id. Sheet IDs are arbitrary numbers that uniquely identify a sheet. Most users won't need this.
#### ```const worksheet xlnt::workbook::sheet_by_id(std::size_t id) const```
Returns the worksheet with a sheetId of id. Sheet IDs are arbitrary numbers that uniquely identify a sheet. Most users won't need this.
#### ```bool xlnt::workbook::contains(const std::string &title) const```
Returns true if this workbook contains a sheet with the given title.
#### ```std::size_t xlnt::workbook::index(worksheet worksheet)```
Returns the index of the given worksheet. The worksheet must be owned by this workbook.
#### ```void xlnt::workbook::remove_sheet(worksheet worksheet)```
Removes the given worksheet from this workbook.
#### ```void xlnt::workbook::clear()```
Sets the contents of this workbook to be equivalent to that of a workbook returned by workbook::empty().
#### ```iterator xlnt::workbook::begin()```
Returns an iterator to the first worksheet in this workbook.
#### ```iterator xlnt::workbook::end()```
Returns an iterator to the worksheet following the last worksheet of the workbook. This worksheet acts as a placeholder; attempting to access it will cause an exception to be thrown.
#### ```const_iterator xlnt::workbook::begin() const```
Returns a const iterator to the first worksheet in this workbook.
#### ```const_iterator xlnt::workbook::end() const```
Returns a const iterator to the worksheet following the last worksheet of the workbook. This worksheet acts as a placeholder; attempting to access it will cause an exception to be thrown.
#### ```const_iterator xlnt::workbook::cbegin() const```
Returns an iterator to the first worksheet in this workbook.
#### ```const_iterator xlnt::workbook::cend() const```
Returns a const iterator to the worksheet following the last worksheet of the workbook. This worksheet acts as a placeholder; attempting to access it will cause an exception to be thrown.
#### ```void xlnt::workbook::apply_to_cells(std::function< void(cell)> f)```
Applies the function "f" to every non-empty cell in every worksheet in this workbook.
#### ```std::vector<std::string> xlnt::workbook::sheet_titles() const```
Returns a temporary vector containing the titles of each sheet in the order of the sheets in the workbook.
#### ```std::size_t xlnt::workbook::sheet_count() const```
Returns the number of sheets in this workbook.
#### ```bool xlnt::workbook::has_core_property(xlnt::core_property type) const```
Returns true if the workbook has the core property with the given name.
#### ```std::vector<xlnt::core_property> xlnt::workbook::core_properties() const```
Returns a vector of the type of each core property that is set to a particular value in this workbook.
#### ```variant xlnt::workbook::core_property(xlnt::core_property type) const```
Returns the value of the given core property.
#### ```void xlnt::workbook::core_property(xlnt::core_property type, const variant &value)```
Sets the given core property to the provided value.
#### ```bool xlnt::workbook::has_extended_property(xlnt::extended_property type) const```
Returns true if the workbook has the extended property with the given name.
#### ```std::vector<xlnt::extended_property> xlnt::workbook::extended_properties() const```
Returns a vector of the type of each extended property that is set to a particular value in this workbook.
#### ```variant xlnt::workbook::extended_property(xlnt::extended_property type) const```
Returns the value of the given extended property.
#### ```void xlnt::workbook::extended_property(xlnt::extended_property type, const variant &value)```
Sets the given extended property to the provided value.
#### ```bool xlnt::workbook::has_custom_property(const std::string &property_name) const```
Returns true if the workbook has the custom property with the given name.
#### ```std::vector<std::string> xlnt::workbook::custom_properties() const```
Returns a vector of the name of each custom property that is set to a particular value in this workbook.
#### ```variant xlnt::workbook::custom_property(const std::string &property_name) const```
Returns the value of the given custom property.
#### ```void xlnt::workbook::custom_property(const std::string &property_name, const variant &value)```
Creates a new custom property in this workbook and sets it to the provided value.
#### ```calendar xlnt::workbook::base_date() const```
Returns the base date used by this workbook. This will generally be windows_1900 except on Apple based systems when it will default to mac_1904 unless otherwise set via void workbook::base_date(calendar base_date).
#### ```void xlnt::workbook::base_date(calendar base_date)```
Sets the base date style of this workbook. This is the date and time that a numeric value of 0 represents.
#### ```bool xlnt::workbook::has_title() const```
Returns true if this workbook has had its title set.
#### ```std::string xlnt::workbook::title() const```
Returns the title of this workbook.
#### ```void xlnt::workbook::title(const std::string &title)```
Sets the title of this workbook to title.
#### ```std::vector<xlnt::named_range> xlnt::workbook::named_ranges() const```
Returns a vector of the named ranges in this workbook.
#### ```void xlnt::workbook::create_named_range(const std::string &name, worksheet worksheet, const range_reference &reference)```
Creates a new names range.
#### ```void xlnt::workbook::create_named_range(const std::string &name, worksheet worksheet, const std::string &reference_string)```
Creates a new names range.
#### ```bool xlnt::workbook::has_named_range(const std::string &name) const```
Returns true if a named range of the given name exists in the workbook.
#### ```class range xlnt::workbook::named_range(const std::string &name)```
Returns the named range with the given name.
#### ```void xlnt::workbook::remove_named_range(const std::string &name)```
Deletes the named range with the given name.
#### ```void xlnt::workbook::save(std::vector< std::uint8_t > &data) const```
Serializes the workbook into an XLSX file and saves the bytes into byte vector data.
#### ```void xlnt::workbook::save(std::vector< std::uint8_t > &data, const std::string &password) const```
Serializes the workbook into an XLSX file encrypted with the given password and saves the bytes into byte vector data.
#### ```void xlnt::workbook::save(const std::string &filename) const```
Serializes the workbook into an XLSX file and saves the data into a file named filename.
#### ```void xlnt::workbook::save(const std::string &filename, const std::string &password) const```
Serializes the workbook into an XLSX file encrypted with the given password and loads the bytes into a file named filename.
#### ```void xlnt::workbook::save(const xlnt::path &filename) const```
Serializes the workbook into an XLSX file and saves the data into a file named filename.
#### ```void xlnt::workbook::save(const xlnt::path &filename, const std::string &password) const```
Serializes the workbook into an XLSX file encrypted with the given password and loads the bytes into a file named filename.
#### ```void xlnt::workbook::save(std::ostream &stream) const```
Serializes the workbook into an XLSX file and saves the data into stream.
#### ```void xlnt::workbook::save(std::ostream &stream, const std::string &password) const```
Serializes the workbook into an XLSX file encrypted with the given password and loads the bytes into the given stream.
#### ```void xlnt::workbook::load(const std::vector< std::uint8_t > &data)```
Interprets byte vector data as an XLSX file and sets the content of this workbook to match that file.
#### ```void xlnt::workbook::load(const std::vector< std::uint8_t > &data, const std::string &password)```
Interprets byte vector data as an XLSX file encrypted with the given password and sets the content of this workbook to match that file.
#### ```void xlnt::workbook::load(const std::string &filename)```
Interprets file with the given filename as an XLSX file and sets the content of this workbook to match that file.
#### ```void xlnt::workbook::load(const std::string &filename, const std::string &password)```
Interprets file with the given filename as an XLSX file encrypted with the given password and sets the content of this workbook to match that file.
#### ```void xlnt::workbook::load(const xlnt::path &filename)```
Interprets file with the given filename as an XLSX file and sets the content of this workbook to match that file.
#### ```void xlnt::workbook::load(const xlnt::path &filename, const std::string &password)```
Interprets file with the given filename as an XLSX file encrypted with the given password and sets the content of this workbook to match that file.
#### ```void xlnt::workbook::load(std::istream &stream)```
Interprets data in stream as an XLSX file and sets the content of this workbook to match that file.
#### ```void xlnt::workbook::load(std::istream &stream, const std::string &password)```
Interprets data in stream as an XLSX file encrypted with the given password and sets the content of this workbook to match that file.
#### ```bool xlnt::workbook::has_view() const```
Returns true if this workbook has a view.
#### ```workbook_view xlnt::workbook::view() const```
Returns the view.
#### ```void xlnt::workbook::view(const workbook_view &view)```
Sets the view to view.
#### ```bool xlnt::workbook::has_code_name() const```
Returns true if a code name has been set for this workbook.
#### ```std::string xlnt::workbook::code_name() const```
Returns the code name that was set for this workbook.
#### ```void xlnt::workbook::code_name(const std::string &code_name)```
Sets the code name of this workbook to code_name.
#### ```bool xlnt::workbook::has_file_version() const```
Returns true if this workbook has a file version.
#### ```std::string xlnt::workbook::app_name() const```
Returns the AppName workbook file property.
#### ```std::size_t xlnt::workbook::last_edited() const```
Returns the LastEdited workbook file property.
#### ```std::size_t xlnt::workbook::lowest_edited() const```
Returns the LowestEdited workbook file property.
#### ```std::size_t xlnt::workbook::rup_build() const```
Returns the RupBuild workbook file property.
#### ```bool xlnt::workbook::has_theme() const```
Returns true if this workbook has a theme defined.
#### ```const xlnt::theme& xlnt::workbook::theme() const```
Returns a const reference to this workbook's theme.
#### ```void xlnt::workbook::theme(const class theme &value)```
Sets the theme to value.
#### ```xlnt::format xlnt::workbook::format(std::size_t format_index)```
Returns the cell format at the given index. The index is the position of the format in xl/styles.xml.
#### ```const xlnt::format xlnt::workbook::format(std::size_t format_index) const```
Returns the cell format at the given index. The index is the position of the format in xl/styles.xml.
#### ```xlnt::format xlnt::workbook::create_format(bool default_format=false)```
Creates a new format and returns it.
#### ```void xlnt::workbook::clear_formats()```
Clear all cell-level formatting and formats from the styelsheet. This leaves all other styling in place (e.g. named styles).
#### ```bool xlnt::workbook::has_style(const std::string &name) const```
Returns true if this workbook has a style with a name of name.
#### ```class style xlnt::workbook::style(const std::string &name)```
Returns the named style with the given name.
#### ```const class style xlnt::workbook::style(const std::string &name) const```
Returns the named style with the given name.
#### ```class style xlnt::workbook::create_style(const std::string &name)```
Creates a new style and returns it.
#### ```class style xlnt::workbook::create_builtin_style(std::size_t builtin_id)```
Creates a new style and returns it.
#### ```void xlnt::workbook::clear_styles()```
Clear all named styles from cells and remove the styles from from the styelsheet. This leaves all other styling in place (e.g. cell formats).
#### ```class manifest& xlnt::workbook::manifest()```
Returns a reference to the workbook's internal manifest.
#### ```const class manifest& xlnt::workbook::manifest() const```
Returns a reference to the workbook's internal manifest.
#### ```void xlnt::workbook::add_shared_string(const rich_text &shared, bool allow_duplicates=false)```
Append a shared string to the shared string collection in this workbook. This should not generally be called unless you know what you're doing. If allow_duplicates is false and the string is already in the collection, it will not be added.
#### ```std::vector<rich_text>& xlnt::workbook::shared_strings()```
Returns a reference to the shared strings being used by cells in this workbook.
#### ```const std::vector<rich_text>& xlnt::workbook::shared_strings() const```
Returns a reference to the shared strings being used by cells in this workbook.
#### ```void xlnt::workbook::thumbnail(const std::vector< std::uint8_t > &thumbnail, const std::string &extension, const std::string &content_type)```
Sets the workbook's thumbnail to the given vector of bytes, thumbnail, with the given extension (e.g. jpg) and content_type (e.g. image/jpeg).
#### ```const std::vector<std::uint8_t>& xlnt::workbook::thumbnail() const```
Returns a vector of bytes representing the workbook's thumbnail.
#### ```bool xlnt::workbook::has_calculation_properties() const```
Returns true if this workbook has any calculation properties set.
#### ```class calculation_properties xlnt::workbook::calculation_properties() const```
Returns the calculation properties used in this workbook.
#### ```void xlnt::workbook::calculation_properties(const class calculation_properties &props)```
Sets the calculation properties of this workbook to props.
#### ```workbook& xlnt::workbook::operator=(workbook other)```
Set the contents of this workbook to be equal to those of "other". Other is passed as value to allow for copy-swap idiom.
#### ```worksheet xlnt::workbook::operator[](const std::string &name)```
Return the worksheet with a title of "name".
#### ```worksheet xlnt::workbook::operator[](std::size_t index)```
Return the worksheet at "index".
#### ```bool xlnt::workbook::operator==(const workbook &rhs) const```
Return true if this workbook internal implementation points to the same memory as rhs's.
#### ```bool xlnt::workbook::operator!=(const workbook &rhs) const```
Return true if this workbook internal implementation doesn't point to the same memory as rhs's.
### workbook_view
#### ```bool xlnt::workbook_view::auto_filter_date_groupingundefined```
If true, dates will be grouped when presenting the user with filtering options.
#### ```bool xlnt::workbook_view::minimizedundefined```
If true, the view will be minimized.
#### ```bool xlnt::workbook_view::show_horizontal_scrollundefined```
If true, the horizontal scroll bar will be displayed.
#### ```bool xlnt::workbook_view::show_sheet_tabsundefined```
If true, the sheet tabs will be displayed.
#### ```bool xlnt::workbook_view::show_vertical_scrollundefined```
If true, the vertical scroll bar will be displayed.
#### ```bool xlnt::workbook_view::visibleundefined```
If true, the workbook window will be visible.
#### ```optional<std::size_t> xlnt::workbook_view::active_tabundefined```
The optional index to the active sheet in this view.
#### ```optional<std::size_t> xlnt::workbook_view::first_sheetundefined```
The optional index to the first sheet in this view.
#### ```optional<std::size_t> xlnt::workbook_view::tab_ratioundefined```
The optional ratio between the tabs bar and the horizontal scroll bar.
#### ```optional<std::size_t> xlnt::workbook_view::window_widthundefined```
The width of the workbook window in twips.
#### ```optional<std::size_t> xlnt::workbook_view::window_heightundefined```
The height of the workbook window in twips.
#### ```optional<std::size_t> xlnt::workbook_view::x_windowundefined```
The distance of the workbook window from the left side of the screen in twips.
#### ```optional<std::size_t> xlnt::workbook_view::y_windowundefined```
The distance of the workbook window from the top of the screen in twips.
### worksheet_iterator
#### ```xlnt::worksheet_iterator::worksheet_iterator(workbook &wb, std::size_t index)```
Constructs a worksheet iterator from a workbook and sheet index.
#### ```xlnt::worksheet_iterator::worksheet_iterator(const worksheet_iterator &)```
Copy constructs a worksheet iterator from another iterator.
#### ```worksheet_iterator& xlnt::worksheet_iterator::operator=(const worksheet_iterator &)```
Assigns the iterator so that it points to the same worksheet in the same workbook.
#### ```worksheet xlnt::worksheet_iterator::operator*()```
Dereferences the iterator to return the worksheet it is pointing to. If the iterator points to one-past-the-end of the workbook, an invalid_parameter exception will be thrown.
#### ```bool xlnt::worksheet_iterator::operator==(const worksheet_iterator &comparand) const```
Returns true if this iterator points to the same worksheet as comparand.
#### ```bool xlnt::worksheet_iterator::operator!=(const worksheet_iterator &comparand) const```
Returns true if this iterator doesn't point to the same worksheet as comparand.
#### ```worksheet_iterator xlnt::worksheet_iterator::operator++(int)```
Post-increment the iterator's internal workseet index. Returns a copy of the iterator as it was before being incremented.
#### ```worksheet_iterator& xlnt::worksheet_iterator::operator++()```
Pre-increment the iterator's internal workseet index. Returns a refernce to the same iterator.
### const_worksheet_iterator
#### ```xlnt::const_worksheet_iterator::const_worksheet_iterator(const workbook &wb, std::size_t index)```
Constructs a worksheet iterator from a workbook and sheet index.
#### ```xlnt::const_worksheet_iterator::const_worksheet_iterator(const const_worksheet_iterator &)```
Copy constructs a worksheet iterator from another iterator.
#### ```const_worksheet_iterator& xlnt::const_worksheet_iterator::operator=(const const_worksheet_iterator &)```
Assigns the iterator so that it points to the same worksheet in the same workbook.
#### ```const worksheet xlnt::const_worksheet_iterator::operator*()```
Dereferences the iterator to return the worksheet it is pointing to. If the iterator points to one-past-the-end of the workbook, an invalid_parameter exception will be thrown.
#### ```bool xlnt::const_worksheet_iterator::operator==(const const_worksheet_iterator &comparand) const```
Returns true if this iterator points to the same worksheet as comparand.
#### ```bool xlnt::const_worksheet_iterator::operator!=(const const_worksheet_iterator &comparand) const```
Returns true if this iterator doesn't point to the same worksheet as comparand.
#### ```const_worksheet_iterator xlnt::const_worksheet_iterator::operator++(int)```
Post-increment the iterator's internal workseet index. Returns a copy of the iterator as it was before being incremented.
#### ```const_worksheet_iterator& xlnt::const_worksheet_iterator::operator++()```
Pre-increment the iterator's internal workseet index. Returns a refernce to the same iterator.
## Worksheet Module
### cell_iterator
#### ```xlnt::cell_iterator::cell_iterator(worksheet ws, const cell_reference &start_cell, const range_reference &limits, major_order order, bool skip_null, bool wrap)```
Constructs a cell_iterator from a worksheet, range, and iteration settings.
#### ```xlnt::cell_iterator::cell_iterator(const cell_iterator &other)```
Constructs a cell_iterator as a copy of an existing cell_iterator.
#### ```cell_iterator& xlnt::cell_iterator::operator=(const cell_iterator &rhs)=default```
Assigns this iterator to match the data in rhs.
#### ```cell xlnt::cell_iterator::operator*()```
Dereferences this iterator to return the cell it points to.
#### ```bool xlnt::cell_iterator::operator==(const cell_iterator &other) const```
Returns true if this iterator is equivalent to other.
#### ```bool xlnt::cell_iterator::operator!=(const cell_iterator &other) const```
Returns true if this iterator isn't equivalent to other.
#### ```cell_iterator& xlnt::cell_iterator::operator--()```
Pre-decrements the iterator to point to the previous cell and returns a reference to the iterator.
#### ```cell_iterator xlnt::cell_iterator::operator--(int)```
Post-decrements the iterator to point to the previous cell and return a copy of the iterator before the decrement.
#### ```cell_iterator& xlnt::cell_iterator::operator++()```
Pre-increments the iterator to point to the previous cell and returns a reference to the iterator.
#### ```cell_iterator xlnt::cell_iterator::operator++(int)```
Post-increments the iterator to point to the previous cell and return a copy of the iterator before the decrement.
### const_cell_iterator
#### ```xlnt::const_cell_iterator::const_cell_iterator(worksheet ws, const cell_reference &start_cell, const range_reference &limits, major_order order, bool skip_null, bool wrap)```
Constructs a cell_iterator from a worksheet, range, and iteration settings.
#### ```xlnt::const_cell_iterator::const_cell_iterator(const const_cell_iterator &other)```
Constructs a cell_iterator as a copy of an existing cell_iterator.
#### ```const_cell_iterator& xlnt::const_cell_iterator::operator=(const const_cell_iterator &)=default```
Assigns this iterator to match the data in rhs.
#### ```const cell xlnt::const_cell_iterator::operator*() const```
Dereferences this iterator to return the cell it points to.
#### ```bool xlnt::const_cell_iterator::operator==(const const_cell_iterator &other) const```
Returns true if this iterator is equivalent to other.
#### ```bool xlnt::const_cell_iterator::operator!=(const const_cell_iterator &other) const```
Returns true if this iterator isn't equivalent to other.
#### ```const_cell_iterator& xlnt::const_cell_iterator::operator--()```
Pre-decrements the iterator to point to the previous cell and returns a reference to the iterator.
#### ```const_cell_iterator xlnt::const_cell_iterator::operator--(int)```
Post-decrements the iterator to point to the previous cell and return a copy of the iterator before the decrement.
#### ```const_cell_iterator& xlnt::const_cell_iterator::operator++()```
Pre-increments the iterator to point to the previous cell and returns a reference to the iterator.
#### ```const_cell_iterator xlnt::const_cell_iterator::operator++(int)```
Post-increments the iterator to point to the previous cell and return a copy of the iterator before the decrement.
### cell_vector
#### ```using xlnt::cell_vector::iterator =  cell_iteratorundefined```
Iterate over cells in a cell_vector with an iterator of this type.
#### ```using xlnt::cell_vector::const_iterator =  const_cell_iteratorundefined```
Iterate over const cells in a const cell_vector with an iterator of this type.
#### ```using xlnt::cell_vector::reverse_iterator =  std::reverse_iterator<iterator>undefined```
Iterate over cells in a cell_vector in reverse oreder with an iterator of this type.
#### ```using xlnt::cell_vector::const_reverse_iterator =  std::reverse_iterator<const_iterator>undefined```
Iterate over const cells in a const cell_vector in reverse order with an iterator of this type.
#### ```xlnt::cell_vector::cell_vector(worksheet ws, const cell_reference &cursor, const range_reference &ref, major_order order, bool skip_null, bool wrap)```
Constructs a cell vector pointing to a given range in a given worksheet. order determines whether this vector is a row or a column. If skip_null is true, iterating over this vector will skip empty cells.
#### ```bool xlnt::cell_vector::empty() const```
Returns true if every cell in this vector is null (i.e. the cell doesn't exist in the worksheet).
#### ```cell xlnt::cell_vector::front()```
Returns the first cell in this vector.
#### ```const cell xlnt::cell_vector::front() const```
Returns the first cell in this vector.
#### ```cell xlnt::cell_vector::back()```
Returns the last cell in this vector.
#### ```const cell xlnt::cell_vector::back() const```
Returns the last cell in this vector.
#### ```std::size_t xlnt::cell_vector::length() const```
Returns the distance between the first and last cells in this vector.
#### ```iterator xlnt::cell_vector::begin()```
Returns an iterator to the first cell in this vector.
#### ```iterator xlnt::cell_vector::end()```
Returns an iterator to a cell one-past-the-end of this vector.
#### ```const_iterator xlnt::cell_vector::begin() const```
Returns an iterator to the first cell in this vector.
#### ```const_iterator xlnt::cell_vector::cbegin() const```
Returns an iterator to the first cell in this vector.
#### ```const_iterator xlnt::cell_vector::end() const```
Returns an iterator to a cell one-past-the-end of this vector.
#### ```const_iterator xlnt::cell_vector::cend() const```
Returns an iterator to a cell one-past-the-end of this vector.
#### ```reverse_iterator xlnt::cell_vector::rbegin()```
Returns a reverse iterator to the first cell of this reversed vector.
#### ```reverse_iterator xlnt::cell_vector::rend()```
Returns a reverse iterator to to a cell one-past-the-end of this reversed vector.
#### ```const_reverse_iterator xlnt::cell_vector::rbegin() const```
Returns a reverse iterator to the first cell of this reversed vector.
#### ```const_reverse_iterator xlnt::cell_vector::rend() const```
Returns a reverse iterator to to a cell one-past-the-end of this reversed vector.
#### ```const_reverse_iterator xlnt::cell_vector::crbegin() const```
Returns a reverse iterator to the first cell of this reversed vector.
#### ```const_reverse_iterator xlnt::cell_vector::crend() const```
Returns a reverse iterator to to a cell one-past-the-end of this reversed vector.
#### ```cell xlnt::cell_vector::operator[](std::size_t column_index)```
Returns the cell column_index distance away from the first cell in this vector.
#### ```const cell xlnt::cell_vector::operator[](std::size_t column_index) const```
Returns the cell column_index distance away from the first cell in this vector.
### column_properties
#### ```optional<double> xlnt::column_properties::widthundefined```
The optional width of the column
#### ```bool xlnt::column_properties::custom_widthundefined```
If true, this is a custom width
#### ```optional<std::size_t> xlnt::column_properties::styleundefined```
The style index of this column. This shouldn't be used since style indices aren't supposed to be used directly in xlnt. (TODO)
#### ```bool xlnt::column_properties::hiddenundefined```
If true, this column will be hidden
### header_footer
#### ```undefinedundefined```
Enumerates the three possible locations of a header or footer.
#### ```bool xlnt::header_footer::has_header() const```
True if any text has been added for a header at any location on any page.
#### ```bool xlnt::header_footer::has_footer() const```
True if any text has been added for a footer at any location on any page.
#### ```bool xlnt::header_footer::align_with_margins() const```
True if headers and footers should align to the page margins.
#### ```header_footer& xlnt::header_footer::align_with_margins(bool align)```
Set to true if headers and footers should align to the page margins. Set to false if headers and footers should align to the edge of the page.
#### ```bool xlnt::header_footer::different_odd_even() const```
True if headers and footers differ based on page number.
#### ```bool xlnt::header_footer::different_first() const```
True if headers and footers are different on the first page.
#### ```bool xlnt::header_footer::scale_with_doc() const```
True if headers and footers should scale to match the worksheet.
#### ```header_footer& xlnt::header_footer::scale_with_doc(bool scale)```
Set to true if headers and footers should scale to match the worksheet.
#### ```bool xlnt::header_footer::has_header(location where) const```
True if any text has been added at the given location on any page.
#### ```void xlnt::header_footer::clear_header()```
Remove all headers from all pages.
#### ```void xlnt::header_footer::clear_header(location where)```
Remove header at the given location on any page.
#### ```header_footer& xlnt::header_footer::header(location where, const std::string &text)```
Add a header at the given location with the given text.
#### ```header_footer& xlnt::header_footer::header(location where, const rich_text &text)```
Add a header at the given location with the given text.
#### ```rich_text xlnt::header_footer::header(location where) const```
Get the text of the header at the given location. If headers are different on odd and even pages, the odd header will be returned.
#### ```bool xlnt::header_footer::has_first_page_header() const```
True if a header has been set for the first page at any location.
#### ```bool xlnt::header_footer::has_first_page_header(location where) const```
True if a header has been set for the first page at the given location.
#### ```void xlnt::header_footer::clear_first_page_header()```
Remove all headers from the first page.
#### ```void xlnt::header_footer::clear_first_page_header(location where)```
Remove header from the first page at the given location.
#### ```header_footer& xlnt::header_footer::first_page_header(location where, const rich_text &text)```
Add a header on the first page at the given location with the given text.
#### ```rich_text xlnt::header_footer::first_page_header(location where) const```
Get the text of the first page header at the given location. If no first page header has been set, the general header for that location will be returned.
#### ```bool xlnt::header_footer::has_odd_even_header() const```
True if different headers have been set for odd and even pages.
#### ```bool xlnt::header_footer::has_odd_even_header(location where) const```
True if different headers have been set for odd and even pages at the given location.
#### ```void xlnt::header_footer::clear_odd_even_header()```
Remove odd/even headers at all locations.
#### ```void xlnt::header_footer::clear_odd_even_header(location where)```
Remove odd/even headers at the given location.
#### ```header_footer& xlnt::header_footer::odd_even_header(location where, const rich_text &odd, const rich_text &even)```
Add a header for odd pages at the given location with the given text.
#### ```rich_text xlnt::header_footer::odd_header(location where) const```
Get the text of the odd page header at the given location. If no odd page header has been set, the general header for that location will be returned.
#### ```rich_text xlnt::header_footer::even_header(location where) const```
Get the text of the even page header at the given location. If no even page header has been set, the general header for that location will be returned.
#### ```bool xlnt::header_footer::has_footer(location where) const```
True if any text has been added at the given location on any page.
#### ```void xlnt::header_footer::clear_footer()```
Remove all footers from all pages.
#### ```void xlnt::header_footer::clear_footer(location where)```
Remove footer at the given location on any page.
#### ```header_footer& xlnt::header_footer::footer(location where, const std::string &text)```
Add a footer at the given location with the given text.
#### ```header_footer& xlnt::header_footer::footer(location where, const rich_text &text)```
Add a footer at the given location with the given text.
#### ```rich_text xlnt::header_footer::footer(location where) const```
Get the text of the footer at the given location. If footers are different on odd and even pages, the odd footer will be returned.
#### ```bool xlnt::header_footer::has_first_page_footer() const```
True if a footer has been set for the first page at any location.
#### ```bool xlnt::header_footer::has_first_page_footer(location where) const```
True if a footer has been set for the first page at the given location.
#### ```void xlnt::header_footer::clear_first_page_footer()```
Remove all footers from the first page.
#### ```void xlnt::header_footer::clear_first_page_footer(location where)```
Remove footer from the first page at the given location.
#### ```header_footer& xlnt::header_footer::first_page_footer(location where, const rich_text &text)```
Add a footer on the first page at the given location with the given text.
#### ```rich_text xlnt::header_footer::first_page_footer(location where) const```
Get the text of the first page footer at the given location. If no first page footer has been set, the general footer for that location will be returned.
#### ```bool xlnt::header_footer::has_odd_even_footer() const```
True if different footers have been set for odd and even pages.
#### ```bool xlnt::header_footer::has_odd_even_footer(location where) const```
True if different footers have been set for odd and even pages at the given location.
#### ```void xlnt::header_footer::clear_odd_even_footer()```
Remove odd/even footers at all locations.
#### ```void xlnt::header_footer::clear_odd_even_footer(location where)```
Remove odd/even footers at the given location.
#### ```header_footer& xlnt::header_footer::odd_even_footer(location where, const rich_text &odd, const rich_text &even)```
Add a footer for odd pages at the given location with the given text.
#### ```rich_text xlnt::header_footer::odd_footer(location where) const```
Get the text of the odd page footer at the given location. If no odd page footer has been set, the general footer for that location will be returned.
#### ```rich_text xlnt::header_footer::even_footer(location where) const```
Get the text of the even page footer at the given location. If no even page footer has been set, the general footer for that location will be returned.
### page_margins
#### ```xlnt::page_margins::page_margins()```
Constructs a page margins objects with Excel-default margins.
#### ```double xlnt::page_margins::top() const```
Returns the top margin
#### ```void xlnt::page_margins::top(double top)```
Sets the top margin to top
#### ```double xlnt::page_margins::left() const```
Returns the left margin
#### ```void xlnt::page_margins::left(double left)```
Sets the left margin to left
#### ```double xlnt::page_margins::bottom() const```
Returns the bottom margin
#### ```void xlnt::page_margins::bottom(double bottom)```
Sets the bottom margin to bottom
#### ```double xlnt::page_margins::right() const```
Returns the right margin
#### ```void xlnt::page_margins::right(double right)```
Sets the right margin to right
#### ```double xlnt::page_margins::header() const```
Returns the header margin
#### ```void xlnt::page_margins::header(double header)```
Sets the header margin to header
#### ```double xlnt::page_margins::footer() const```
Returns the footer margin
#### ```void xlnt::page_margins::footer(double footer)```
Sets the footer margin to footer
### page_setup
#### ```xlnt::page_setup::page_setup()```
Default constructor.
#### ```xlnt::page_break xlnt::page_setup::page_break() const```
Returns the page break.
#### ```void xlnt::page_setup::page_break(xlnt::page_break b)```
Sets the page break to b.
#### ```xlnt::sheet_state xlnt::page_setup::sheet_state() const```
Returns the current sheet state of this page setup.
#### ```void xlnt::page_setup::sheet_state(xlnt::sheet_state sheet_state)```
Sets the sheet state to sheet_state.
#### ```xlnt::paper_size xlnt::page_setup::paper_size() const```
Returns the paper size which should be used to print the worksheet using this page setup.
#### ```void xlnt::page_setup::paper_size(xlnt::paper_size paper_size)```
Sets the paper size of this page setup.
#### ```xlnt::orientation xlnt::page_setup::orientation() const```
Returns the orientation of the worksheet using this page setup.
#### ```void xlnt::page_setup::orientation(xlnt::orientation orientation)```
Sets the orientation of the page.
#### ```bool xlnt::page_setup::fit_to_page() const```
Returns true if this worksheet should be scaled to fit on a single page during printing.
#### ```void xlnt::page_setup::fit_to_page(bool fit_to_page)```
If true, forces the worksheet to be scaled to fit on a single page during printing.
#### ```bool xlnt::page_setup::fit_to_height() const```
Returns true if the height of this worksheet should be scaled to fit on a printed page.
#### ```void xlnt::page_setup::fit_to_height(bool fit_to_height)```
Sets whether the height of the page should be scaled to fit on a printed page.
#### ```bool xlnt::page_setup::fit_to_width() const```
Returns true if the width of this worksheet should be scaled to fit on a printed page.
#### ```void xlnt::page_setup::fit_to_width(bool fit_to_width)```
Sets whether the width of the page should be scaled to fit on a printed page.
#### ```void xlnt::page_setup::horizontal_centered(bool horizontal_centered)```
Sets whether the worksheet should be centered horizontall on the page if it takes up less than a full page.
#### ```bool xlnt::page_setup::horizontal_centered() const```
Returns whether horizontal centering has been enabled.
#### ```void xlnt::page_setup::vertical_centered(bool vertical_centered)```
Sets whether the worksheet should be vertically centered on the page if it takes up less than a full page.
#### ```bool xlnt::page_setup::vertical_centered() const```
Returns whether vertical centering has been enabled.
#### ```void xlnt::page_setup::scale(double scale)```
Sets the factor by which the page should be scaled during printing.
#### ```double xlnt::page_setup::scale() const```
Returns the factor by which the page should be scaled during printing.
### pane
#### ```optional<cell_reference> xlnt::pane::top_left_cellundefined```
The optional top left cell
#### ```pane_state xlnt::pane::stateundefined```
The state of the pane
#### ```pane_corner xlnt::pane::active_paneundefined```
The pane which contains the active cell
#### ```row_t xlnt::pane::y_splitundefined```
The row where the split should take place
#### ```column_t xlnt::pane::x_splitundefined```
The column where the split should take place
#### ```bool xlnt::pane::operator==(const pane &rhs) const```
Returns true if this pane is equal to rhs based on its top-left cell, state, active pane, and x/y split location.
### range
#### ```using xlnt::range::iterator =  range_iteratorundefined```
Alias for the iterator type
#### ```using xlnt::range::const_iterator =  const_range_iteratorundefined```
Alias for the const iterator type
#### ```using xlnt::range::reverse_iterator =  std::reverse_iterator<iterator>undefined```
Alias for the reverse iterator type
#### ```using xlnt::range::const_reverse_iterator =  std::reverse_iterator<const_iterator>undefined```
Alias for the const reverse iterator type
#### ```xlnt::range::range(worksheet ws, const range_reference &reference, major_order order=major_order::row, bool skip_null=false)```
Constructs a range on the given worksheet.
#### ```xlnt::range::~range()```
Desctructor
#### ```xlnt::range::range(const range &)=default```
Default copy constructor.
#### ```cell_vector xlnt::range::vector(std::size_t n)```
Returns a vector pointing to the n-th row or column in this range (depending on major order).
#### ```const cell_vector xlnt::range::vector(std::size_t n) const```
Returns a vector pointing to the n-th row or column in this range (depending on major order).
#### ```class cell xlnt::range::cell(const cell_reference &ref)```
Returns a cell in the range relative to its top left cell.
#### ```const class cell xlnt::range::cell(const cell_reference &ref) const```
Returns a cell in the range relative to its top left cell.
#### ```range_reference xlnt::range::reference() const```
Returns the reference defining the bounds of this range.
#### ```std::size_t xlnt::range::length() const```
Returns the number of rows or columns in this range (depending on major order).
#### ```bool xlnt::range::contains(const cell_reference &ref)```
Returns true if the given cell exists in the parent worksheet of this range.
#### ```range xlnt::range::alignment(const xlnt::alignment &new_alignment)```
Sets the alignment of all cells in the range to new_alignment and returns the range.
#### ```range xlnt::range::border(const xlnt::border &new_border)```
Sets the border of all cells in the range to new_border and returns the range.
#### ```range xlnt::range::fill(const xlnt::fill &new_fill)```
Sets the fill of all cells in the range to new_fill and returns the range.
#### ```range xlnt::range::font(const xlnt::font &new_font)```
Sets the font of all cells in the range to new_font and returns the range.
#### ```range xlnt::range::number_format(const xlnt::number_format &new_number_format)```
Sets the number format of all cells in the range to new_number_format and returns the range.
#### ```range xlnt::range::protection(const xlnt::protection &new_protection)```
Sets the protection of all cells in the range to new_protection and returns the range.
#### ```range xlnt::range::style(const class style &new_style)```
Sets the named style applied to all cells in this range to a style named style_name.
#### ```range xlnt::range::style(const std::string &style_name)```
Sets the named style applied to all cells in this range to a style named style_name. If this style has not been previously created in the workbook, a key_not_found exception will be thrown.
#### ```cell_vector xlnt::range::front()```
Returns the first row or column in this range.
#### ```const cell_vector xlnt::range::front() const```
Returns the first row or column in this range.
#### ```cell_vector xlnt::range::back()```
Returns the last row or column in this range.
#### ```const cell_vector xlnt::range::back() const```
Returns the last row or column in this range.
#### ```iterator xlnt::range::begin()```
Returns an iterator to the first row or column in this range.
#### ```iterator xlnt::range::end()```
Returns an iterator to one past the last row or column in this range.
#### ```const_iterator xlnt::range::begin() const```
Returns an iterator to the first row or column in this range.
#### ```const_iterator xlnt::range::end() const```
Returns an iterator to one past the last row or column in this range.
#### ```const_iterator xlnt::range::cbegin() const```
Returns an iterator to the first row or column in this range.
#### ```const_iterator xlnt::range::cend() const```
Returns an iterator to one past the last row or column in this range.
#### ```reverse_iterator xlnt::range::rbegin()```
Returns a reverse iterator to the first row or column in this range.
#### ```reverse_iterator xlnt::range::rend()```
Returns a reverse iterator to one past the last row or column in this range.
#### ```const_reverse_iterator xlnt::range::rbegin() const```
Returns a reverse iterator to the first row or column in this range.
#### ```const_reverse_iterator xlnt::range::rend() const```
Returns a reverse iterator to one past the last row or column in this range.
#### ```const_reverse_iterator xlnt::range::crbegin() const```
Returns a reverse iterator to the first row or column in this range.
#### ```const_reverse_iterator xlnt::range::crend() const```
Returns a reverse iterator to one past the last row or column in this range.
#### ```void xlnt::range::apply(std::function< void(class cell)> f)```
Applies function f to all cells in the range
#### ```cell_vector xlnt::range::operator[](std::size_t n)```
Returns the n-th row or column in this range.
#### ```const cell_vector xlnt::range::operator[](std::size_t n) const```
Returns the n-th row or column in this range.
#### ```bool xlnt::range::operator==(const range &comparand) const```
Returns true if this range is equivalent to comparand.
#### ```bool xlnt::range::operator!=(const range &comparand) const```
Returns true if this range is not equivalent to comparand.
### range_iterator
#### ```xlnt::range_iterator::range_iterator(worksheet &ws, const cell_reference &cursor, const range_reference &bounds, major_order order, bool skip_null)```
Constructs a range iterator on a worksheet, cell pointing to the current row or column, range bounds, an order, and whether or not to skip null column/rows.
#### ```xlnt::range_iterator::range_iterator(const range_iterator &other)```
Copy constructor.
#### ```cell_vector xlnt::range_iterator::operator*() const```
Dereference the iterator to return a column or row.
#### ```range_iterator& xlnt::range_iterator::operator=(const range_iterator &)=default```
Default assignment operator.
#### ```bool xlnt::range_iterator::operator==(const range_iterator &other) const```
Returns true if this iterator is equivalent to other.
#### ```bool xlnt::range_iterator::operator!=(const range_iterator &other) const```
Returns true if this iterator is not equivalent to other.
#### ```range_iterator& xlnt::range_iterator::operator--()```
Pre-decrement the iterator to point to the previous row/column.
#### ```range_iterator xlnt::range_iterator::operator--(int)```
Post-decrement the iterator to point to the previous row/column.
#### ```range_iterator& xlnt::range_iterator::operator++()```
Pre-increment the iterator to point to the next row/column.
#### ```range_iterator xlnt::range_iterator::operator++(int)```
Post-increment the iterator to point to the next row/column.
### const_range_iterator
#### ```xlnt::const_range_iterator::const_range_iterator(const worksheet &ws, const cell_reference &cursor, const range_reference &bounds, major_order order, bool skip_null)```
Constructs a range iterator on a worksheet, cell pointing to the current row or column, range bounds, an order, and whether or not to skip null column/rows.
#### ```xlnt::const_range_iterator::const_range_iterator(const const_range_iterator &other)```
Copy constructor.
#### ```const cell_vector xlnt::const_range_iterator::operator*() const```
Dereferennce the iterator to return the current column/row.
#### ```const_range_iterator& xlnt::const_range_iterator::operator=(const const_range_iterator &)=default```
Default assignment operator.
#### ```bool xlnt::const_range_iterator::operator==(const const_range_iterator &other) const```
Returns true if this iterator is equivalent to other.
#### ```bool xlnt::const_range_iterator::operator!=(const const_range_iterator &other) const```
Returns true if this iterator is not equivalent to other.
#### ```const_range_iterator& xlnt::const_range_iterator::operator--()```
Pre-decrement the iterator to point to the next row/column.
#### ```const_range_iterator xlnt::const_range_iterator::operator--(int)```
Post-decrement the iterator to point to the next row/column.
#### ```const_range_iterator& xlnt::const_range_iterator::operator++()```
Pre-increment the iterator to point to the next row/column.
#### ```const_range_iterator xlnt::const_range_iterator::operator++(int)```
Post-increment the iterator to point to the next row/column.
### range_reference
#### ```static range_reference xlnt::range_reference::make_absolute(const range_reference &relative_reference)```
Converts relative reference coordinates to absolute coordinates (B12 -> $B$12)
#### ```xlnt::range_reference::range_reference()```
Constructs a range reference equal to A1:A1
#### ```xlnt::range_reference::range_reference(const std::string &range_string)```
Constructs a range reference equivalent to the provided range_string in the form <top_left>:<bottom_right>.
#### ```xlnt::range_reference::range_reference(const char *range_string)```
Constructs a range reference equivalent to the provided range_string in the form <top_left>:<bottom_right>.
#### ```xlnt::range_reference::range_reference(const std::pair< cell_reference, cell_reference > &reference_pair)```
Constructs a range reference from a pair of cell references.
#### ```xlnt::range_reference::range_reference(const cell_reference &start, const cell_reference &end)```
Constructs a range reference from cell references indicating top left and bottom right coordinates of the range.
#### ```xlnt::range_reference::range_reference(column_t column_index_start, row_t row_index_start, column_t column_index_end, row_t row_index_end)```
Constructs a range reference from column and row indices.
#### ```bool xlnt::range_reference::is_single_cell() const```
Returns true if the range has a width and height of 1 cell.
#### ```std::size_t xlnt::range_reference::width() const```
Returns the number of columns encompassed by this range.
#### ```std::size_t xlnt::range_reference::height() const```
Returns the number of rows encompassed by this range.
#### ```cell_reference xlnt::range_reference::top_left() const```
Returns the coordinate of the top left cell of this range.
#### ```cell_reference xlnt::range_reference::top_right() const```
Returns the coordinate of the top right cell of this range.
#### ```cell_reference xlnt::range_reference::bottom_left() const```
Returns the coordinate of the bottom left cell of this range.
#### ```cell_reference xlnt::range_reference::bottom_right() const```
Returns the coordinate of the bottom right cell of this range.
#### ```range_reference xlnt::range_reference::make_offset(int column_offset, int row_offset) const```
Returns a new range reference with the same width and height as this range but shifted by the given number of columns and rows.
#### ```std::string xlnt::range_reference::to_string() const```
Returns a string representation of this range.
#### ```bool xlnt::range_reference::operator==(const range_reference &comparand) const```
Returns true if this range is equivalent to the other range.
#### ```bool xlnt::range_reference::operator==(const std::string &reference_string) const```
Returns true if this range is equivalent to the string representation of the other range.
#### ```bool xlnt::range_reference::operator==(const char *reference_string) const```
Returns true if this range is equivalent to the string representation of the other range.
#### ```bool xlnt::range_reference::operator!=(const range_reference &comparand) const```
Returns true if this range is not equivalent to the other range.
#### ```bool xlnt::range_reference::operator!=(const std::string &reference_string) const```
Returns true if this range is not equivalent to the string representation of the other range.
#### ```bool xlnt::range_reference::operator!=(const char *reference_string) const```
Returns true if this range is not equivalent to the string representation of the other range.
### row_properties
#### ```optional<double> xlnt::row_properties::heightundefined```
Optional height
#### ```bool xlnt::row_properties::custom_heightundefined```
Whether or not the height is different from the default
#### ```bool xlnt::row_properties::hiddenundefined```
Whether or not the row should be hidden
#### ```optional<std::size_t> xlnt::row_properties::styleundefined```
The index to the style used by all cells in this row
### selection
#### ```bool xlnt::selection::has_active_cell() const```
Returns true if this selection has a defined active cell.
#### ```cell_reference xlnt::selection::active_cell() const```
Returns the cell reference of the active cell.
#### ```void xlnt::selection::active_cell(const cell_reference &ref)```
Sets the active cell to that pointed to by ref.
#### ```range_reference xlnt::selection::sqref() const```
Returns the range encompassed by this selection.
#### ```pane_corner xlnt::selection::pane() const```
Returns the sheet quadrant of this selection.
#### ```void xlnt::selection::pane(pane_corner corner)```
Sets the sheet quadrant of this selection to corner.
#### ```bool xlnt::selection::operator==(const selection &rhs) const```
Returns true if this selection is equal to rhs based on its active cell, sqref, and pane.
### sheet_protection
#### ```static std::string xlnt::sheet_protection::hash_password(const std::string &password)```
Calculates and returns the hash of the given protection password.
#### ```void xlnt::sheet_protection::password(const std::string &password)```
Sets the protection password to password.
#### ```std::string xlnt::sheet_protection::hashed_password() const```
Returns the hash of the password set for this sheet protection.
### sheet_view
#### ```void xlnt::sheet_view::id(std::size_t new_id)```
Sets the ID of this view to new_id.
#### ```std::size_t xlnt::sheet_view::id() const```
Returns the ID of this view.
#### ```bool xlnt::sheet_view::has_pane() const```
Returns true if this view has a pane defined.
#### ```struct pane& xlnt::sheet_view::pane()```
Returns a reference to this view's pane.
#### ```const struct pane& xlnt::sheet_view::pane() const```
Returns a reference to this view's pane.
#### ```void xlnt::sheet_view::clear_pane()```
Removes the defined pane from this view.
#### ```void xlnt::sheet_view::pane(const struct pane &new_pane)```
Sets the pane of this view to new_pane.
#### ```bool xlnt::sheet_view::has_selections() const```
Returns true if this view has any selections.
#### ```void xlnt::sheet_view::add_selection(const class selection &new_selection)```
Adds the given selection to the collection of selections.
#### ```void xlnt::sheet_view::clear_selections()```
Removes all selections.
#### ```std::vector<xlnt::selection> xlnt::sheet_view::selections() const```
Returns the collection of selections as a vector.
#### ```class xlnt::selection& xlnt::sheet_view::selection(std::size_t index)```
Returns the selection at the given index.
#### ```void xlnt::sheet_view::show_grid_lines(bool show)```
If show is true, grid lines will be shown for sheets using this view.
#### ```bool xlnt::sheet_view::show_grid_lines() const```
Returns true if grid lines will be shown for sheets using this view.
#### ```void xlnt::sheet_view::default_grid_color(bool is_default)```
If is_default is true, the default grid color will be used.
#### ```bool xlnt::sheet_view::default_grid_color() const```
Returns true if the default grid color will be used.
#### ```void xlnt::sheet_view::type(sheet_view_type new_type)```
Sets the type of this view.
#### ```sheet_view_type xlnt::sheet_view::type() const```
Returns the type of this view.
#### ```bool xlnt::sheet_view::operator==(const sheet_view &rhs) const```
Returns true if this view is requal to rhs based on its id, grid lines setting, default grid color, pane, and selections.
### worksheet
#### ```using xlnt::worksheet::iterator =  range_iteratorundefined```
Iterate over a non-const worksheet with an iterator of this type.
#### ```using xlnt::worksheet::const_iterator =  const_range_iteratorundefined```
Iterate over a non-const worksheet with an iterator of this type.
#### ```using xlnt::worksheet::reverse_iterator =  std::reverse_iterator<iterator>undefined```
Iterate in reverse over a non-const worksheet with an iterator of this type.
#### ```using xlnt::worksheet::const_reverse_iterator =  std::reverse_iterator<const_iterator>undefined```
Iterate in reverse order over a const worksheet with an iterator of this type.
#### ```friend class detail::xlsx_consumerundefined```
#### ```friend class detail::xlsx_producerundefined```
#### ```xlnt::worksheet::worksheet()```
Construct a null worksheet. No methods should be called on such a worksheet.
#### ```xlnt::worksheet::worksheet(const worksheet &rhs)```
Copy constructor. This worksheet will point to the same memory as rhs's worksheet.
#### ```class workbook& xlnt::worksheet::workbook()```
Returns a reference to the workbook this worksheet is owned by.
#### ```const class workbook& xlnt::worksheet::workbook() const```
Returns a reference to the workbook this worksheet is owned by.
#### ```void xlnt::worksheet::garbage_collect()```
Deletes data held in the worksheet that does not affect the internal data or display. For example, unreference styles and empty cells will be removed.
#### ```std::size_t xlnt::worksheet::id() const```
Returns the unique numeric identifier of this worksheet. This will sometimes but not necessarily be the index of the worksheet in the workbook.
#### ```void xlnt::worksheet::id(std::size_t id)```
Set the unique numeric identifier. The id defaults to the lowest unused id in the workbook so this should not be called without a good reason.
#### ```std::string xlnt::worksheet::title() const```
Returns the title of this sheet.
#### ```void xlnt::worksheet::title(const std::string &title)```
Sets the title of this sheet.
#### ```cell_reference xlnt::worksheet::frozen_panes() const```
Returns the top left corner of the region above and to the left of which panes are frozen.
#### ```void xlnt::worksheet::freeze_panes(cell top_left_cell)```
Freeze panes above and to the left of top_left_cell.
#### ```void xlnt::worksheet::freeze_panes(const cell_reference &top_left_coordinate)```
Freeze panes above and to the left of top_left_coordinate.
#### ```void xlnt::worksheet::unfreeze_panes()```
Remove frozen panes. The data in those panes will be unaffectedthis affects only the view.
#### ```bool xlnt::worksheet::has_frozen_panes() const```
Returns true if this sheet has a frozen row, frozen column, or both.
#### ```bool xlnt::worksheet::has_cell(const cell_reference &reference) const```
Returns true if this sheet has an initialized cell at the given reference.
#### ```class cell xlnt::worksheet::cell(const cell_reference &reference)```
Returns the cell at the given reference. If the cell doesn't exist, it will be initialized to null before being returned.
#### ```const class cell xlnt::worksheet::cell(const cell_reference &reference) const```
Returns the cell at the given reference. If the cell doesn't exist, an invalid_parameter exception will be thrown.
#### ```class cell xlnt::worksheet::cell(column_t column, row_t row)```
Returns the cell at the given column and row. If the cell doesn't exist, it will be initialized to null before being returned.
#### ```const class cell xlnt::worksheet::cell(column_t column, row_t row) const```
Returns the cell at the given column and row. If the cell doesn't exist, an invalid_parameter exception will be thrown.
#### ```class range xlnt::worksheet::range(const std::string &reference_string)```
Returns the range defined by reference string. If reference string is the name of a previously-defined named range in the sheet, it will be returned.
#### ```const class range xlnt::worksheet::range(const std::string &reference_string) const```
Returns the range defined by reference string. If reference string is the name of a previously-defined named range in the sheet, it will be returned.
#### ```class range xlnt::worksheet::range(const range_reference &reference)```
Returns the range defined by reference.
#### ```const class range xlnt::worksheet::range(const range_reference &reference) const```
Returns the range defined by reference.
#### ```class range xlnt::worksheet::rows(bool skip_null=true)```
Returns a range encompassing all cells in this sheet which will be iterated upon in row-major order. If skip_null is true (default), empty rows and cells will be skipped during iteration of the range.
#### ```const class range xlnt::worksheet::rows(bool skip_null=true) const```
Returns a range encompassing all cells in this sheet which will be iterated upon in row-major order. If skip_null is true (default), empty rows and cells will be skipped during iteration of the range.
#### ```class range xlnt::worksheet::columns(bool skip_null=true)```
Returns a range ecompassing all cells in this sheet which will be iterated upon in column-major order. If skip_null is true (default), empty columns and cells will be skipped during iteration of the range.
#### ```const class range xlnt::worksheet::columns(bool skip_null=true) const```
Returns a range ecompassing all cells in this sheet which will be iterated upon in column-major order. If skip_null is true (default), empty columns and cells will be skipped during iteration of the range.
#### ```xlnt::column_properties& xlnt::worksheet::column_properties(column_t column)```
Returns the column properties for the given column.
#### ```const xlnt::column_properties& xlnt::worksheet::column_properties(column_t column) const```
Returns the column properties for the given column.
#### ```bool xlnt::worksheet::has_column_properties(column_t column) const```
Returns true if column properties have been set for the given column.
#### ```void xlnt::worksheet::add_column_properties(column_t column, const class column_properties &props)```
Sets column properties for the given column to props.
#### ```double xlnt::worksheet::column_width(column_t column) const```
Calculates the width of the given column. This will be the default column width if a custom width is not set on this column's column_properties.
#### ```xlnt::row_properties& xlnt::worksheet::row_properties(row_t row)```
Returns the row properties for the given row.
#### ```const xlnt::row_properties& xlnt::worksheet::row_properties(row_t row) const```
Returns the row properties for the given row.
#### ```bool xlnt::worksheet::has_row_properties(row_t row) const```
Returns true if row properties have been set for the given row.
#### ```void xlnt::worksheet::add_row_properties(row_t row, const class row_properties &props)```
Sets row properties for the given row to props.
#### ```double xlnt::worksheet::row_height(row_t row) const```
Calculate the height of the given row. This will be the default row height if a custom height is not set on this row's row_properties.
#### ```cell_reference xlnt::worksheet::point_pos(int left, int top) const```
Returns a reference to the cell at the given point coordinates.
#### ```void xlnt::worksheet::create_named_range(const std::string &name, const std::string &reference_string)```
Creates a new named range with the given name encompassing the string representing a range.
#### ```void xlnt::worksheet::create_named_range(const std::string &name, const range_reference &reference)```
Creates a new named range with the given name encompassing the given range reference.
#### ```bool xlnt::worksheet::has_named_range(const std::string &name) const```
Returns true if this worksheet contains a named range with the given name.
#### ```class range xlnt::worksheet::named_range(const std::string &name)```
Returns the named range with the given name. Throws key_not_found exception if the named range doesn't exist.
#### ```const class range xlnt::worksheet::named_range(const std::string &name) const```
Returns the named range with the given name. Throws key_not_found exception if the named range doesn't exist.
#### ```void xlnt::worksheet::remove_named_range(const std::string &name)```
Removes a named range with the given name.
#### ```row_t xlnt::worksheet::lowest_row() const```
Returns the row of the first non-empty cell in the worksheet.
#### ```row_t xlnt::worksheet::highest_row() const```
Returns the row of the last non-empty cell in the worksheet.
#### ```row_t xlnt::worksheet::next_row() const```
Returns the row directly below the last non-empty cell in the worksheet.
#### ```column_t xlnt::worksheet::lowest_column() const```
Returns the column of the first non-empty cell in the worksheet.
#### ```column_t xlnt::worksheet::highest_column() const```
Returns the column of the last non-empty cell in the worksheet.
#### ```range_reference xlnt::worksheet::calculate_dimension() const```
Returns a range_reference pointing to the full range of non-empty cells in the worksheet.
#### ```void xlnt::worksheet::merge_cells(const std::string &reference_string)```
Merges the cells within the range represented by the given string.
#### ```void xlnt::worksheet::merge_cells(const range_reference &reference)```
Merges the cells within the given range.
#### ```void xlnt::worksheet::unmerge_cells(const std::string &reference_string)```
Removes the merging of the cells in the range represented by the given string.
#### ```void xlnt::worksheet::unmerge_cells(const range_reference &reference)```
Removes the merging of the cells in the given range.
#### ```std::vector<range_reference> xlnt::worksheet::merged_ranges() const```
Returns a vector of references of all merged ranges in the worksheet.
#### ```bool xlnt::worksheet::operator==(const worksheet &other) const```
Returns true if this worksheet refers to the same worksheet as other.
#### ```bool xlnt::worksheet::operator!=(const worksheet &other) const```
Returns true if this worksheet doesn't refer to the same worksheet as other.
#### ```bool xlnt::worksheet::operator==(std::nullptr_t) const```
Returns true if this worksheet is null.
#### ```bool xlnt::worksheet::operator!=(std::nullptr_t) const```
Returns true if this worksheet is not null.
#### ```void xlnt::worksheet::operator=(const worksheet &other)```
Sets the internal pointer of this worksheet object to point to other.
#### ```class cell xlnt::worksheet::operator[](const cell_reference &reference)```
Convenience method for worksheet::cell method.
#### ```const class cell xlnt::worksheet::operator[](const cell_reference &reference) const```
Convenience method for worksheet::cell method.
#### ```bool xlnt::worksheet::compare(const worksheet &other, bool reference) const```
Returns true if this worksheet is equal to other. If reference is true, the comparison will only check that both worksheets point to the same sheet in the same workbook.
#### ```bool xlnt::worksheet::has_page_setup() const```
Returns true if this worksheet has a page setup.
#### ```xlnt::page_setup xlnt::worksheet::page_setup() const```
Returns the page setup for this sheet.
#### ```void xlnt::worksheet::page_setup(const struct page_setup &setup)```
Sets the page setup for this sheet to setup.
#### ```bool xlnt::worksheet::has_page_margins() const```
Returns true if this page has margins.
#### ```xlnt::page_margins xlnt::worksheet::page_margins() const```
Returns the margins of this sheet.
#### ```void xlnt::worksheet::page_margins(const class page_margins &margins)```
Sets the margins of this sheet to margins.
#### ```range_reference xlnt::worksheet::auto_filter() const```
Returns the current auto-filter of this sheet.
#### ```void xlnt::worksheet::auto_filter(const std::string &range_string)```
Sets the auto-filter of this sheet to the range defined by range_string.
#### ```void xlnt::worksheet::auto_filter(const xlnt::range &range)```
Sets the auto-filter of this sheet to the given range.
#### ```void xlnt::worksheet::auto_filter(const range_reference &reference)```
Sets the auto-filter of this sheet to the range defined by reference.
#### ```void xlnt::worksheet::clear_auto_filter()```
Clear the auto-filter from this sheet.
#### ```bool xlnt::worksheet::has_auto_filter() const```
Returns true if this sheet has an auto-filter set.
#### ```void xlnt::worksheet::reserve(std::size_t n)```
Reserve n rows. This can be optionally called before adding many rows to improve performance.
#### ```bool xlnt::worksheet::has_header_footer() const```
Returns true if this sheet has a header/footer.
#### ```class header_footer xlnt::worksheet::header_footer() const```
Returns the header/footer of this sheet.
#### ```void xlnt::worksheet::header_footer(const class header_footer &new_header_footer)```
Sets the header/footer of this sheet to new_header_footer.
#### ```xlnt::sheet_state xlnt::worksheet::sheet_state() const```
Returns the sheet state of this sheet.
#### ```void xlnt::worksheet::sheet_state(xlnt::sheet_state state)```
Sets the sheet state of this sheet.
#### ```iterator xlnt::worksheet::begin()```
Returns an iterator to the first row in this sheet.
#### ```iterator xlnt::worksheet::end()```
Returns an iterator one past the last row in this sheet.
#### ```const_iterator xlnt::worksheet::begin() const```
Return a constant iterator to the first row in this sheet.
#### ```const_iterator xlnt::worksheet::end() const```
Returns a constant iterator to one past the last row in this sheet.
#### ```const_iterator xlnt::worksheet::cbegin() const```
Return a constant iterator to the first row in this sheet.
#### ```const_iterator xlnt::worksheet::cend() const```
Returns a constant iterator to one past the last row in this sheet.
#### ```void xlnt::worksheet::print_title_rows(row_t first_row, row_t last_row)```
Sets rows to repeat at top during printing.
#### ```void xlnt::worksheet::print_title_rows(row_t last_row)```
Sets rows to repeat at top during printing.
#### ```void xlnt::worksheet::print_title_cols(column_t first_column, column_t last_column)```
Sets columns to repeat at left during printing.
#### ```void xlnt::worksheet::print_title_cols(column_t last_column)```
Sets columns to repeat at left during printing.
#### ```std::string xlnt::worksheet::print_titles() const```
Returns a string representation of the defined print titles.
#### ```void xlnt::worksheet::print_area(const std::string &print_area)```
Sets the print area of this sheet to print_area.
#### ```range_reference xlnt::worksheet::print_area() const```
Returns the print area defined for this sheet.
#### ```bool xlnt::worksheet::has_view() const```
Returns true if this sheet has any number of views defined.
#### ```sheet_view xlnt::worksheet::view(std::size_t index=0) const```
Returns the view at the given index.
#### ```void xlnt::worksheet::add_view(const sheet_view &new_view)```
Adds new_view to the set of available views for this sheet.
#### ```void xlnt::worksheet::clear_page_breaks()```
Remove all manual column and row page breaks (represented as dashed blue lines in the page view in Excel).
#### ```const std::vector<row_t>& xlnt::worksheet::page_break_rows() const```
Returns vector where each element represents a row which will break a page below it.
#### ```void xlnt::worksheet::page_break_at_row(row_t row)```
Add a page break at the given row.
#### ```const std::vector<column_t>& xlnt::worksheet::page_break_columns() const```
Returns vector where each element represents a column which will break a page to the right.
#### ```void xlnt::worksheet::page_break_at_column(column_t column)```
Add a page break at the given column.
