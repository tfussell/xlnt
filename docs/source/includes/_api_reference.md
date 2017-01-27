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
#### ```void xlnt::cell::value(T value)```
Sets the value of this cell to the given value. Overloads exist for most C++ fundamental types like bool, int, etc. as well as for std::string and xlnt datetime types: date, time, datetime, and timedelta.
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
#### ```void xlnt::cell::hyperlink(const std::string &value)```
Adds a hyperlink to this cell pointing to the URI of the given value.
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
### cell_reference
#### ```static std::pair<std::string, row_t> xlnt::cell_reference::split_reference(const std::string &reference_string)```
Split a coordinate string like "A1" into an equivalent pair like {"A", 1}.
#### ```static std::pair<std::string, row_t> xlnt::cell_reference::split_reference(const std::string &reference_string, bool &absolute_column, bool &absolute_row)```
Split a coordinate string like "A1" into an equivalent pair like {"A", 1}. Reference parameters absolute_column and absolute_row will be set to true if column part or row part are prefixed by a dollar-sign indicating they are absolute, otherwise false.
#### ```xlnt::cell_reference::cell_reference()```
Default constructor makes a reference to the top-left-most cell, "A1".
#### ```xlnt::cell_reference::cell_reference(const char *reference_string)```
Constructs a cell_reference from a string reprenting a cell coordinate (e.g. $B14).
#### ```xlnt::cell_reference::cell_reference(const std::string &reference_string)```
Constructs a cell_reference from a string reprenting a cell coordinate (e.g. $B14).
#### ```xlnt::cell_reference::cell_reference(column_t column, row_t row)```
Constructs a cell_reference from a 1-indexed column index and row index.
#### ```cell_reference& xlnt::cell_reference::make_absolute(bool absolute_column=true, bool absolute_row=true)```
Convert a coordinate to an absolute coordinate string (e.g. B12 -> $B$12) Defaulting to true, absolute_column and absolute_row can optionally control whether the resulting cell_reference has an absolute column (e.g. B12 -> $B12) and absolute row (e.g. B12 -> B$12) respectively.
#### ```bool xlnt::cell_reference::column_absolute() const```
Return true if the reference refers to an absolute column, otherwise false.
#### ```void xlnt::cell_reference::column_absolute(bool absolute_column)```
Make this reference have an absolute column if absolute_column is true, otherwise not absolute.
#### ```bool xlnt::cell_reference::row_absolute() const```
Return true if the reference refers to an absolute row, otherwise false.
#### ```void xlnt::cell_reference::row_absolute(bool absolute_row)```
Make this reference have an absolute row if absolute_row is true, otherwise not absolute.
#### ```column_t xlnt::cell_reference::column() const```
Return a string that identifies the column of this reference (e.g. second column from left is "B")
#### ```void xlnt::cell_reference::column(const std::string &column_string)```
Set the column of this reference from a string that identifies a particular column.
#### ```column_t::index_t xlnt::cell_reference::column_index() const```
Return a 1-indexed numeric index of the column of this reference.
#### ```void xlnt::cell_reference::column_index(column_t column)```
Set the column of this reference from a 1-indexed number that identifies a particular column.
#### ```row_t xlnt::cell_reference::row() const```
Return a 1-indexed numeric index of the row of this reference.
#### ```void xlnt::cell_reference::row(row_t row)```
Set the row of this reference from a 1-indexed number that identifies a particular row.
#### ```cell_reference xlnt::cell_reference::make_offset(int column_offset, int row_offset) const```
Return a cell_reference offset from this cell_reference by the number of columns and rows specified by the parameters. A negative value for column_offset or row_offset results in a reference above or left of this cell_reference, respectively.
#### ```std::string xlnt::cell_reference::to_string() const```
Return a string like "A1" for cell_reference(1, 1).
#### ```range_reference xlnt::cell_reference::to_range() const```
Return a 1x1 range_reference containing only this cell_reference.
#### ```range_reference xlnt::cell_reference::operator,(const cell_reference &other) const```
I've always wanted to overload the comma operator. cell_reference("A", 1), cell_reference("B", 1) will return range_reference(cell_reference("A", 1), cell_reference("B", 1))
#### ```bool xlnt::cell_reference::operator==(const cell_reference &comparand) const```
Return true if this reference is identical to comparand including in absoluteness of column and row.
#### ```bool xlnt::cell_reference::operator==(const std::string &reference_string) const```
Construct a cell_reference from reference_string and return the result of their comparison.
#### ```bool xlnt::cell_reference::operator==(const char *reference_string) const```
Construct a cell_reference from reference_string and return the result of their comparison.
#### ```bool xlnt::cell_reference::operator!=(const cell_reference &comparand) const```
Return true if this reference is not identical to comparand including in absoluteness of column and row.
#### ```bool xlnt::cell_reference::operator!=(const std::string &reference_string) const```
Construct a cell_reference from reference_string and return the result of their comparison.
#### ```bool xlnt::cell_reference::operator!=(const char *reference_string) const```
Construct a cell_reference from reference_string and return the result of their comparison.
### comment
#### ```xlnt::comment::comment()```
Constructs a new blank comment.
#### ```xlnt::comment::comment(const rich_text &text, const std::string &author)```
Constructs a new comment with the given text and author.
#### ```xlnt::comment::comment(const std::string &text, const std::string &author)```
Constructs a new comment with the given unformatted text and author.
#### ```rich_text xlnt::comment::text() const```
Return the text that will be displayed for this comment.
#### ```std::string xlnt::comment::plain_text() const```
Return the plain text that will be displayed for this comment without formatting information.
#### ```std::string xlnt::comment::author() const```
Return the author of this comment.
#### ```void xlnt::comment::hide()```
Make this comment only visible when the associated cell is hovered.
#### ```void xlnt::comment::show()```
Make this comment always visible.
#### ```bool xlnt::comment::visible() const```
Returns true if this comment is not hidden.
#### ```void xlnt::comment::position(int left, int top)```
Set the absolute position of this cell to the given coordinates.
#### ```int xlnt::comment::left() const```
Returns the distance from the left side of the sheet to the left side of the comment.
#### ```int xlnt::comment::top() const```
Returns the distance from the top of the sheet to the top of the comment.
#### ```void xlnt::comment::size(int width, int height)```
Set the size of the comment.
#### ```int xlnt::comment::width() const```
Returns the width of this comment.
#### ```int xlnt::comment::height() const```
Returns the height of this comment.
#### ```bool operator==(const comment &left, const comment &right)```
Return true if both comments are equivalent.
### column_t
#### ```using xlnt::column_t::index_t =  std::uint32_tundefined```
#### ```index_t xlnt::column_t::indexundefined```
Internal numeric value of this column index.
#### ```static index_t xlnt::column_t::column_index_from_string(const std::string &column_string)```
Convert a column letter into a column number (e.g. B -> 2)
#### ```static std::string xlnt::column_t::column_string_from_index(index_t column_index)```
Convert a column number into a column letter (3 -> 'C')
#### ```xlnt::column_t::column_t()```
Default column_t is the first (left-most) column.
#### ```xlnt::column_t::column_t(index_t column_index)```
Construct a column from a number.
#### ```xlnt::column_t::column_t(const std::string &column_string)```
Construct a column from a string.
#### ```xlnt::column_t::column_t(const char *column_string)```
Construct a column from a string.
#### ```xlnt::column_t::column_t(const column_t &other)```
Copy constructor
#### ```xlnt::column_t::column_t(column_t &&other)```
Move constructor
#### ```std::string xlnt::column_t::column_string() const```
Return a string representation of this column index.
#### ```column_t& xlnt::column_t::operator=(column_t rhs)```
Set this column to be the same as rhs's and return reference to self.
#### ```column_t& xlnt::column_t::operator=(const std::string &rhs)```
Set this column to be equal to rhs and return reference to self.
#### ```column_t& xlnt::column_t::operator=(const char *rhs)```
Set this column to be equal to rhs and return reference to self.
#### ```bool xlnt::column_t::operator==(const column_t &other) const```
Return true if this column refers to the same column as other.
#### ```bool xlnt::column_t::operator!=(const column_t &other) const```
Return true if this column doesn't refer to the same column as other.
#### ```bool xlnt::column_t::operator==(int other) const```
Return true if this column refers to the same column as other.
#### ```bool xlnt::column_t::operator==(index_t other) const```
Return true if this column refers to the same column as other.
#### ```bool xlnt::column_t::operator==(const std::string &other) const```
Return true if this column refers to the same column as other.
#### ```bool xlnt::column_t::operator==(const char *other) const```
Return true if this column refers to the same column as other.
#### ```bool xlnt::column_t::operator!=(int other) const```
Return true if this column doesn't refer to the same column as other.
#### ```bool xlnt::column_t::operator!=(index_t other) const```
Return true if this column doesn't refer to the same column as other.
#### ```bool xlnt::column_t::operator!=(const std::string &other) const```
Return true if this column doesn't refer to the same column as other.
#### ```bool xlnt::column_t::operator!=(const char *other) const```
Return true if this column doesn't refer to the same column as other.
#### ```bool xlnt::column_t::operator>(const column_t &other) const```
Return true if other is to the right of this column.
#### ```bool xlnt::column_t::operator>=(const column_t &other) const```
Return true if other is to the right of or equal to this column.
#### ```bool xlnt::column_t::operator<(const column_t &other) const```
Return true if other is to the left of this column.
#### ```bool xlnt::column_t::operator<=(const column_t &other) const```
Return true if other is to the left of or equal to this column.
#### ```bool xlnt::column_t::operator>(const column_t::index_t &other) const```
Return true if other is to the right of this column.
#### ```bool xlnt::column_t::operator>=(const column_t::index_t &other) const```
Return true if other is to the right of or equal to this column.
#### ```bool xlnt::column_t::operator<(const column_t::index_t &other) const```
Return true if other is to the left of this column.
#### ```bool xlnt::column_t::operator<=(const column_t::index_t &other) const```
Return true if other is to the left of or equal to this column.
#### ```column_t& xlnt::column_t::operator++()```
Pre-increment this column, making it point to the column one to the right.
#### ```column_t& xlnt::column_t::operator--()```
Pre-deccrement this column, making it point to the column one to the left.
#### ```column_t xlnt::column_t::operator++(int)```
Post-increment this column, making it point to the column one to the right and returning the old column.
#### ```column_t xlnt::column_t::operator--(int)```
Post-decrement this column, making it point to the column one to the left and returning the old column.
#### ```column_t& xlnt::column_t::operator+=(const column_t &rhs)```
Add rhs to this column and return a reference to this column.
#### ```column_t& xlnt::column_t::operator-=(const column_t &rhs)```
Subtrac rhs from this column and return a reference to this column.
#### ```column_t& xlnt::column_t::operator*=(const column_t &rhs)```
Multiply this column by rhs and return a reference to this column.
#### ```column_t& xlnt::column_t::operator/=(const column_t &rhs)```
Divide this column by rhs and return a reference to this column.
#### ```column_t& xlnt::column_t::operator%=(const column_t &rhs)```
Mod this column by rhs and return a reference to this column.
#### ```column_t operator+(column_t lhs, const column_t &rhs)```
Return the result of adding rhs to this column.
#### ```column_t operator-(column_t lhs, const column_t &rhs)```
Return the result of subtracing lhs by rhs column.
#### ```column_t operator*(column_t lhs, const column_t &rhs)```
Return the result of multiply lhs by rhs column.
#### ```column_t operator/(column_t lhs, const column_t &rhs)```
Return the result of divide lhs by rhs column.
#### ```column_t operator%(column_t lhs, const column_t &rhs)```
Return the result of mod lhs by rhs column.
#### ```bool operator>(const column_t::index_t &left, const column_t &right)```
Return true if other is to the right of this column.
#### ```bool operator>=(const column_t::index_t &left, const column_t &right)```
Return true if other is to the right of or equal to this column.
#### ```bool operator<(const column_t::index_t &left, const column_t &right)```
Return true if other is to the left of this column.
#### ```bool operator<=(const column_t::index_t &left, const column_t &right)```
Return true if other is to the left of or equal to this column.
#### ```void swap(column_t &left, column_t &right)```
Swap the columns that left and right refer to.
### column_hash
#### ```std::size_t xlnt::column_hash::operator()(const column_t &k) const```
### column_t >
#### ```size_t std::hash< xlnt::column_t >::operator()(const xlnt::column_t &k) const```
### rich_text
#### ```xlnt::rich_text::rich_text()=default```
#### ```xlnt::rich_text::rich_text(const std::string &plain_text)```
#### ```xlnt::rich_text::rich_text(const std::string &plain_text, const class font &text_font)```
#### ```xlnt::rich_text::rich_text(const rich_text_run &single_run)```
#### ```void xlnt::rich_text::clear()```
Remove all text runs from this text.
#### ```void xlnt::rich_text::plain_text(const std::string &s)```
Clear any runs in this text and add a single run with default formatting and the given string as its textual content.
#### ```std::string xlnt::rich_text::plain_text() const```
Combine the textual content of each text run in order and return the result.
#### ```std::vector<rich_text_run> xlnt::rich_text::runs() const```
Returns a copy of the individual runs that comprise this text.
#### ```void xlnt::rich_text::runs(const std::vector< rich_text_run > &new_runs)```
Set the runs of this text all at once.
#### ```void xlnt::rich_text::add_run(const rich_text_run &t)```
Add a new run to the end of the set of runs.
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
#### ```std::string xlnt::manifest::register_relationship(const class relationship &rel)```
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
#### ```xlnt::relationship::relationship(const std::string &id, relationship_type t, const uri &source, const uri &target, xlnt::target_mode mode)```
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
#### ```xlnt::uri::uri(const uri &base, const uri &relative)```
#### ```xlnt::uri::uri(const uri &base, const path &relative)```
#### ```xlnt::uri::uri(const std::string &uri_string)```
#### ```bool xlnt::uri::is_relative() const```
#### ```bool xlnt::uri::is_absolute() const```
#### ```std::string xlnt::uri::scheme() const```
#### ```std::string xlnt::uri::authority() const```
#### ```bool xlnt::uri::has_authentication() const```
#### ```std::string xlnt::uri::authentication() const```
#### ```std::string xlnt::uri::username() const```
#### ```std::string xlnt::uri::password() const```
#### ```std::string xlnt::uri::host() const```
#### ```bool xlnt::uri::has_port() const```
#### ```std::size_t xlnt::uri::port() const```
#### ```class path xlnt::uri::path() const```
#### ```bool xlnt::uri::has_query() const```
#### ```std::string xlnt::uri::query() const```
#### ```bool xlnt::uri::has_fragment() const```
#### ```std::string xlnt::uri::fragment() const```
#### ```std::string xlnt::uri::to_string() const```
#### ```uri xlnt::uri::make_absolute(const uri &base)```
#### ```uri xlnt::uri::make_reference(const uri &base)```
#### ```bool operator==(const uri &left, const uri &right)```
## Styles Module
### alignment
#### ```bool xlnt::alignment::shrink() const```
#### ```alignment& xlnt::alignment::shrink(bool shrink_to_fit)```
#### ```bool xlnt::alignment::wrap() const```
#### ```alignment& xlnt::alignment::wrap(bool wrap_text)```
#### ```optional<int> xlnt::alignment::indent() const```
#### ```alignment& xlnt::alignment::indent(int indent_size)```
#### ```optional<int> xlnt::alignment::rotation() const```
#### ```alignment& xlnt::alignment::rotation(int text_rotation)```
#### ```optional<horizontal_alignment> xlnt::alignment::horizontal() const```
#### ```alignment& xlnt::alignment::horizontal(horizontal_alignment horizontal)```
#### ```optional<vertical_alignment> xlnt::alignment::vertical() const```
#### ```alignment& xlnt::alignment::vertical(vertical_alignment vertical)```
#### ```bool operator==(const alignment &left, const alignment &right)```
Returns true if left is exactly equal to right.
#### ```bool operator!=(const alignment &left, const alignment &right)```
Returns true if left is not exactly equal to right.
### border
#### ```static const std::vector<border_side>& xlnt::border::all_sides()```
#### ```xlnt::border::border()```
#### ```optional<border_property> xlnt::border::side(border_side s) const```
#### ```border& xlnt::border::side(border_side s, const border_property &prop)```
#### ```optional<diagonal_direction> xlnt::border::diagonal() const```
#### ```border& xlnt::border::diagonal(diagonal_direction dir)```
#### ```bool operator==(const border &left, const border &right)```
Returns true if left is exactly equal to right.
#### ```bool operator!=(const border &left, const border &right)```
Returns true if left is not exactly equal to right.
### border_property
#### ```optional<class color> xlnt::border::border_property::color() const```
#### ```border_property& xlnt::border::border_property::color(const xlnt::color &c)```
#### ```optional<border_style> xlnt::border::border_property::style() const```
#### ```border_property& xlnt::border::border_property::style(border_style style)```
#### ```bool operator==(const border_property &left, const border_property &right)```
Returns true if left is exactly equal to right.
#### ```bool operator!=(const border_property &left, const border_property &right)```
Returns true if left is not exactly equal to right.
### indexed_color
#### ```xlnt::indexed_color::indexed_color(std::size_t index)```
#### ```std::size_t xlnt::indexed_color::index() const```
#### ```void xlnt::indexed_color::index(std::size_t index)```
### theme_color
#### ```xlnt::theme_color::theme_color(std::size_t index)```
#### ```std::size_t xlnt::theme_color::index() const```
#### ```void xlnt::theme_color::index(std::size_t index)```
### rgb_color
#### ```xlnt::rgb_color::rgb_color(const std::string &hex_string)```
#### ```xlnt::rgb_color::rgb_color(std::uint8_t r, std::uint8_t g, std::uint8_t b, std::uint8_t a=255)```
#### ```std::string xlnt::rgb_color::hex_string() const```
#### ```std::uint8_t xlnt::rgb_color::red() const```
#### ```std::uint8_t xlnt::rgb_color::green() const```
#### ```std::uint8_t xlnt::rgb_color::blue() const```
#### ```std::uint8_t xlnt::rgb_color::alpha() const```
#### ```std::array<std::uint8_t, 3> xlnt::rgb_color::rgb() const```
#### ```std::array<std::uint8_t, 4> xlnt::rgb_color::rgba() const```
### color
#### ```static const color xlnt::color::black()```
#### ```static const color xlnt::color::white()```
#### ```static const color xlnt::color::red()```
#### ```static const color xlnt::color::darkred()```
#### ```static const color xlnt::color::blue()```
#### ```static const color xlnt::color::darkblue()```
#### ```static const color xlnt::color::green()```
#### ```static const color xlnt::color::darkgreen()```
#### ```static const color xlnt::color::yellow()```
#### ```static const color xlnt::color::darkyellow()```
#### ```xlnt::color::color()```
#### ```xlnt::color::color(const rgb_color &rgb)```
#### ```xlnt::color::color(const indexed_color &indexed)```
#### ```xlnt::color::color(const theme_color &theme)```
#### ```color_type xlnt::color::type() const```
#### ```bool xlnt::color::is_auto() const```
#### ```void xlnt::color::auto_(bool value)```
#### ```const rgb_color& xlnt::color::rgb() const```
#### ```const indexed_color& xlnt::color::indexed() const```
#### ```const theme_color& xlnt::color::theme() const```
#### ```double xlnt::color::tint() const```
#### ```void xlnt::color::tint(double tint)```
#### ```bool xlnt::color::operator==(const color &other) const```
#### ```bool xlnt::color::operator!=(const color &other) const```
### pattern_fill
#### ```xlnt::pattern_fill::pattern_fill()```
#### ```pattern_fill_type xlnt::pattern_fill::type() const```
#### ```pattern_fill& xlnt::pattern_fill::type(pattern_fill_type new_type)```
#### ```optional<color> xlnt::pattern_fill::foreground() const```
#### ```pattern_fill& xlnt::pattern_fill::foreground(const color &foreground)```
#### ```optional<color> xlnt::pattern_fill::background() const```
#### ```pattern_fill& xlnt::pattern_fill::background(const color &background)```
#### ```bool operator==(const pattern_fill &left, const pattern_fill &right)```
Returns true if left is exactly equal to right.
#### ```bool operator!=(const pattern_fill &left, const pattern_fill &right)```
Returns true if left is not exactly equal to right.
### gradient_fill
#### ```xlnt::gradient_fill::gradient_fill()```
#### ```gradient_fill_type xlnt::gradient_fill::type() const```
#### ```gradient_fill& xlnt::gradient_fill::type(gradient_fill_type new_type)```
#### ```gradient_fill& xlnt::gradient_fill::degree(double degree)```
#### ```double xlnt::gradient_fill::degree() const```
#### ```double xlnt::gradient_fill::left() const```
#### ```gradient_fill& xlnt::gradient_fill::left(double value)```
#### ```double xlnt::gradient_fill::right() const```
#### ```gradient_fill& xlnt::gradient_fill::right(double value)```
#### ```double xlnt::gradient_fill::top() const```
#### ```gradient_fill& xlnt::gradient_fill::top(double value)```
#### ```double xlnt::gradient_fill::bottom() const```
#### ```gradient_fill& xlnt::gradient_fill::bottom(double value)```
#### ```gradient_fill& xlnt::gradient_fill::add_stop(double position, color stop_color)```
#### ```gradient_fill& xlnt::gradient_fill::clear_stops()```
#### ```std::unordered_map<double, color> xlnt::gradient_fill::stops() const```
#### ```bool operator==(const gradient_fill &left, const gradient_fill &right)```
Returns true if left is exactly equal to right.
#### ```bool operator!=(const gradient_fill &left, const gradient_fill &right)```
Returns true if left is not exactly equal to right.
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
#### ```bool operator==(const fill &left, const fill &right)```
Returns true if left is exactly equal to right.
#### ```bool operator!=(const fill &left, const fill &right)```
Returns true if left is not exactly equal to right.
### font
#### ```undefinedundefined```
#### ```bool operator==(const font &left, const font &right)```
Returns true if left is exactly equal to right.
#### ```bool operator!=(const font &left, const font &right)```
Returns true if left is not exactly equal to right.
#### ```xlnt::font::font()```
#### ```font& xlnt::font::bold(bool bold)```
#### ```bool xlnt::font::bold() const```
#### ```font& xlnt::font::superscript(bool superscript)```
#### ```bool xlnt::font::superscript() const```
#### ```font& xlnt::font::italic(bool italic)```
#### ```bool xlnt::font::italic() const```
#### ```font& xlnt::font::strikethrough(bool strikethrough)```
#### ```bool xlnt::font::strikethrough() const```
#### ```font& xlnt::font::underline(underline_style new_underline)```
#### ```bool xlnt::font::underlined() const```
#### ```underline_style xlnt::font::underline() const```
#### ```bool xlnt::font::has_size() const```
#### ```font& xlnt::font::size(double size)```
#### ```double xlnt::font::size() const```
#### ```bool xlnt::font::has_name() const```
#### ```font& xlnt::font::name(const std::string &name)```
#### ```std::string xlnt::font::name() const```
#### ```bool xlnt::font::has_color() const```
#### ```font& xlnt::font::color(const color &c)```
#### ```xlnt::color xlnt::font::color() const```
#### ```bool xlnt::font::has_family() const```
#### ```font& xlnt::font::family(std::size_t family)```
#### ```std::size_t xlnt::font::family() const```
#### ```bool xlnt::font::has_scheme() const```
#### ```font& xlnt::font::scheme(const std::string &scheme)```
#### ```std::string xlnt::font::scheme() const```
### format
#### ```friend struct detail::stylesheetundefined```
#### ```std::size_t xlnt::format::id() const```
#### ```class alignment& xlnt::format::alignment()```
#### ```const class alignment& xlnt::format::alignment() const```
#### ```format xlnt::format::alignment(const xlnt::alignment &new_alignment, bool applied)```
#### ```bool xlnt::format::alignment_applied() const```
#### ```class border& xlnt::format::border()```
#### ```const class border& xlnt::format::border() const```
#### ```format xlnt::format::border(const xlnt::border &new_border, bool applied)```
#### ```bool xlnt::format::border_applied() const```
#### ```class fill& xlnt::format::fill()```
#### ```const class fill& xlnt::format::fill() const```
#### ```format xlnt::format::fill(const xlnt::fill &new_fill, bool applied)```
#### ```bool xlnt::format::fill_applied() const```
#### ```class font& xlnt::format::font()```
#### ```const class font& xlnt::format::font() const```
#### ```format xlnt::format::font(const xlnt::font &new_font, bool applied)```
#### ```bool xlnt::format::font_applied() const```
#### ```class number_format& xlnt::format::number_format()```
#### ```const class number_format& xlnt::format::number_format() const```
#### ```format xlnt::format::number_format(const xlnt::number_format &new_number_format, bool applied)```
#### ```bool xlnt::format::number_format_applied() const```
#### ```class protection& xlnt::format::protection()```
#### ```const class protection& xlnt::format::protection() const```
#### ```format xlnt::format::protection(const xlnt::protection &new_protection, bool applied)```
#### ```bool xlnt::format::protection_applied() const```
#### ```bool xlnt::format::has_style() const```
#### ```void xlnt::format::clear_style()```
#### ```format xlnt::format::style(const std::string &name)```
#### ```format xlnt::format::style(const class style &new_style)```
#### ```const class style xlnt::format::style() const```
### number_format
#### ```static const number_format xlnt::number_format::general()```
#### ```static const number_format xlnt::number_format::text()```
#### ```static const number_format xlnt::number_format::number()```
#### ```static const number_format xlnt::number_format::number_00()```
#### ```static const number_format xlnt::number_format::number_comma_separated1()```
#### ```static const number_format xlnt::number_format::percentage()```
#### ```static const number_format xlnt::number_format::percentage_00()```
#### ```static const number_format xlnt::number_format::date_yyyymmdd2()```
#### ```static const number_format xlnt::number_format::date_yymmdd()```
#### ```static const number_format xlnt::number_format::date_ddmmyyyy()```
#### ```static const number_format xlnt::number_format::date_dmyslash()```
#### ```static const number_format xlnt::number_format::date_dmyminus()```
#### ```static const number_format xlnt::number_format::date_dmminus()```
#### ```static const number_format xlnt::number_format::date_myminus()```
#### ```static const number_format xlnt::number_format::date_xlsx14()```
#### ```static const number_format xlnt::number_format::date_xlsx15()```
#### ```static const number_format xlnt::number_format::date_xlsx16()```
#### ```static const number_format xlnt::number_format::date_xlsx17()```
#### ```static const number_format xlnt::number_format::date_xlsx22()```
#### ```static const number_format xlnt::number_format::date_datetime()```
#### ```static const number_format xlnt::number_format::date_time1()```
#### ```static const number_format xlnt::number_format::date_time2()```
#### ```static const number_format xlnt::number_format::date_time3()```
#### ```static const number_format xlnt::number_format::date_time4()```
#### ```static const number_format xlnt::number_format::date_time5()```
#### ```static const number_format xlnt::number_format::date_time6()```
#### ```static bool xlnt::number_format::is_builtin_format(std::size_t builtin_id)```
#### ```static const number_format& xlnt::number_format::from_builtin_id(std::size_t builtin_id)```
#### ```xlnt::number_format::number_format()```
#### ```xlnt::number_format::number_format(std::size_t builtin_id)```
#### ```xlnt::number_format::number_format(const std::string &code)```
#### ```xlnt::number_format::number_format(const std::string &code, std::size_t custom_id)```
#### ```void xlnt::number_format::format_string(const std::string &format_code)```
#### ```void xlnt::number_format::format_string(const std::string &format_code, std::size_t custom_id)```
#### ```std::string xlnt::number_format::format_string() const```
#### ```bool xlnt::number_format::has_id() const```
#### ```void xlnt::number_format::id(std::size_t id)```
#### ```std::size_t xlnt::number_format::id() const```
#### ```std::string xlnt::number_format::format(const std::string &text) const```
#### ```std::string xlnt::number_format::format(long double number, calendar base_date) const```
#### ```bool xlnt::number_format::is_date_format() const```
#### ```bool operator==(const number_format &left, const number_format &right)```
Returns true if left is exactly equal to right.
#### ```bool operator!=(const number_format &left, const number_format &right)```
Returns true if left is not exactly equal to right.
### protection
#### ```static protection xlnt::protection::unlocked_and_visible()```
#### ```static protection xlnt::protection::locked_and_visible()```
#### ```static protection xlnt::protection::unlocked_and_hidden()```
#### ```static protection xlnt::protection::locked_and_hidden()```
#### ```xlnt::protection::protection()```
#### ```bool xlnt::protection::locked() const```
#### ```protection& xlnt::protection::locked(bool locked)```
#### ```bool xlnt::protection::hidden() const```
#### ```protection& xlnt::protection::hidden(bool hidden)```
#### ```bool operator==(const protection &left, const protection &right)```
Returns true if left is exactly equal to right.
#### ```bool operator!=(const protection &left, const protection &right)```
Returns true if left is not exactly equal to right.
### style
#### ```friend struct detail::stylesheetundefined```
#### ```friend class detail::xlsx_consumerundefined```
#### ```xlnt::style::style()=delete```
Delete zero-argument constructor
#### ```xlnt::style::style(const style &other)=default```
Default copy constructor
#### ```std::string xlnt::style::name() const```
Return the name of this style.
#### ```style xlnt::style::name(const std::string &name)```
#### ```bool xlnt::style::hidden() const```
#### ```style xlnt::style::hidden(bool value)```
#### ```optional<bool> xlnt::style::custom() const```
#### ```style xlnt::style::custom(bool value)```
#### ```optional<std::size_t> xlnt::style::builtin_id() const```
#### ```style xlnt::style::builtin_id(std::size_t builtin_id)```
#### ```class alignment& xlnt::style::alignment()```
#### ```const class alignment& xlnt::style::alignment() const```
#### ```style xlnt::style::alignment(const xlnt::alignment &new_alignment, bool applied=true)```
#### ```bool xlnt::style::alignment_applied() const```
#### ```class border& xlnt::style::border()```
#### ```const class border& xlnt::style::border() const```
#### ```style xlnt::style::border(const xlnt::border &new_border, bool applied=true)```
#### ```bool xlnt::style::border_applied() const```
#### ```class fill& xlnt::style::fill()```
#### ```const class fill& xlnt::style::fill() const```
#### ```style xlnt::style::fill(const xlnt::fill &new_fill, bool applied=true)```
#### ```bool xlnt::style::fill_applied() const```
#### ```class font& xlnt::style::font()```
#### ```const class font& xlnt::style::font() const```
#### ```style xlnt::style::font(const xlnt::font &new_font, bool applied=true)```
#### ```bool xlnt::style::font_applied() const```
#### ```class number_format& xlnt::style::number_format()```
#### ```const class number_format& xlnt::style::number_format() const```
#### ```style xlnt::style::number_format(const xlnt::number_format &new_number_format, bool applied=true)```
#### ```bool xlnt::style::number_format_applied() const```
#### ```class protection& xlnt::style::protection()```
#### ```const class protection& xlnt::style::protection() const```
#### ```style xlnt::style::protection(const xlnt::protection &new_protection, bool applied=true)```
#### ```bool xlnt::style::protection_applied() const```
#### ```bool xlnt::style::operator==(const style &other) const```
## Utils Module
### date
#### ```int xlnt::date::yearundefined```
#### ```int xlnt::date::monthundefined```
#### ```int xlnt::date::dayundefined```
#### ```static date xlnt::date::today()```
Return the current date according to the system time.
#### ```static date xlnt::date::from_number(int days_since_base_year, calendar base_date)```
Return a date by adding days_since_base_year to base_date. This includes leap years.
#### ```xlnt::date::date(int year_, int month_, int day_)```
#### ```int xlnt::date::to_number(calendar base_date) const```
Return the number of days between this date and base_date.
#### ```int xlnt::date::weekday() const```
#### ```bool xlnt::date::operator==(const date &comparand) const```
Return true if this date is equal to comparand.
### datetime
#### ```int xlnt::datetime::yearundefined```
#### ```int xlnt::datetime::monthundefined```
#### ```int xlnt::datetime::dayundefined```
#### ```int xlnt::datetime::hourundefined```
#### ```int xlnt::datetime::minuteundefined```
#### ```int xlnt::datetime::secondundefined```
#### ```int xlnt::datetime::microsecondundefined```
#### ```static datetime xlnt::datetime::now()```
Return the current date and time according to the system time.
#### ```static datetime xlnt::datetime::today()```
Return the current date and time according to the system time. This is equivalent to datetime::now().
#### ```static datetime xlnt::datetime::from_number(long double number, calendar base_date)```
Return a datetime from number by converting the integer part into a date and the fractional part into a time according to date::from_number and time::from_number.
#### ```static datetime xlnt::datetime::from_iso_string(const std::string &iso_string)```
#### ```xlnt::datetime::datetime(const date &d, const time &t)```
#### ```xlnt::datetime::datetime(int year_, int month_, int day_, int hour_=0, int minute_=0, int second_=0, int microsecond_=0)```
#### ```std::string xlnt::datetime::to_string() const```
#### ```std::string xlnt::datetime::to_iso_string() const```
#### ```long double xlnt::datetime::to_number(calendar base_date) const```
#### ```bool xlnt::datetime::operator==(const datetime &comparand) const```
#### ```int xlnt::datetime::weekday() const```
### exception
#### ```xlnt::exception::exception(const std::string &message)```
#### ```xlnt::exception::exception(const exception &)=default```
#### ```virtual xlnt::exception::~exception()```
#### ```void xlnt::exception::message(const std::string &message)```
### invalid_parameter
#### ```xlnt::invalid_parameter::invalid_parameter()```
#### ```xlnt::invalid_parameter::invalid_parameter(const invalid_parameter &)=default```
#### ```virtual xlnt::invalid_parameter::~invalid_parameter()```
### invalid_sheet_title
#### ```xlnt::invalid_sheet_title::invalid_sheet_title(const std::string &title)```
#### ```xlnt::invalid_sheet_title::invalid_sheet_title(const invalid_sheet_title &)=default```
#### ```virtual xlnt::invalid_sheet_title::~invalid_sheet_title()```
### missing_number_format
#### ```xlnt::missing_number_format::missing_number_format()```
#### ```virtual xlnt::missing_number_format::~missing_number_format()```
### invalid_file
#### ```xlnt::invalid_file::invalid_file(const std::string &filename)```
#### ```xlnt::invalid_file::invalid_file(const invalid_file &)=default```
#### ```virtual xlnt::invalid_file::~invalid_file()```
### illegal_character
#### ```xlnt::illegal_character::illegal_character(char c)```
#### ```xlnt::illegal_character::illegal_character(const illegal_character &)=default```
#### ```virtual xlnt::illegal_character::~illegal_character()```
### invalid_data_type
#### ```xlnt::invalid_data_type::invalid_data_type()```
#### ```xlnt::invalid_data_type::invalid_data_type(const invalid_data_type &)=default```
#### ```virtual xlnt::invalid_data_type::~invalid_data_type()```
### invalid_column_string_index
#### ```xlnt::invalid_column_string_index::invalid_column_string_index()```
#### ```xlnt::invalid_column_string_index::invalid_column_string_index(const invalid_column_string_index &)=default```
#### ```virtual xlnt::invalid_column_string_index::~invalid_column_string_index()```
### invalid_cell_reference
#### ```xlnt::invalid_cell_reference::invalid_cell_reference(column_t column, row_t row)```
#### ```xlnt::invalid_cell_reference::invalid_cell_reference(const std::string &reference_string)```
#### ```xlnt::invalid_cell_reference::invalid_cell_reference(const invalid_cell_reference &)=default```
#### ```virtual xlnt::invalid_cell_reference::~invalid_cell_reference()```
### invalid_attribute
#### ```xlnt::invalid_attribute::invalid_attribute()```
#### ```xlnt::invalid_attribute::invalid_attribute(const invalid_attribute &)=default```
#### ```virtual xlnt::invalid_attribute::~invalid_attribute()```
### key_not_found
#### ```xlnt::key_not_found::key_not_found()```
#### ```xlnt::key_not_found::key_not_found(const key_not_found &)=default```
#### ```virtual xlnt::key_not_found::~key_not_found()```
### no_visible_worksheets
#### ```xlnt::no_visible_worksheets::no_visible_worksheets()```
#### ```xlnt::no_visible_worksheets::no_visible_worksheets(const no_visible_worksheets &)=default```
#### ```virtual xlnt::no_visible_worksheets::~no_visible_worksheets()```
### unhandled_switch_case
#### ```xlnt::unhandled_switch_case::unhandled_switch_case()```
#### ```virtual xlnt::unhandled_switch_case::~unhandled_switch_case()```
### unsupported
#### ```xlnt::unsupported::unsupported(const std::string &message)```
#### ```xlnt::unsupported::unsupported(const unsupported &)=default```
#### ```virtual xlnt::unsupported::~unsupported()```
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
#### ```bool operator==(const path &left, const path &right)```
Returns true if left path is equal to right path.
### path >
#### ```size_t std::hash< xlnt::path >::operator()(const xlnt::path &p) const```
### scoped_enum_hash
#### ```std::size_t xlnt::scoped_enum_hash< Enum >::operator()(Enum e) const```
### time
#### ```int xlnt::time::hourundefined```
#### ```int xlnt::time::minuteundefined```
#### ```int xlnt::time::secondundefined```
#### ```int xlnt::time::microsecondundefined```
#### ```static time xlnt::time::now()```
Return the current time according to the system time.
#### ```static time xlnt::time::from_number(long double number)```
Return a time from a number representing a fraction of a day. The integer part of number will be ignored. 0.5 would return time(12, 0, 0, 0) or noon, halfway through the day.
#### ```xlnt::time::time(int hour_=0, int minute_=0, int second_=0, int microsecond_=0)```
#### ```xlnt::time::time(const std::string &time_string)```
#### ```long double xlnt::time::to_number() const```
#### ```bool xlnt::time::operator==(const time &comparand) const```
### timedelta
#### ```int xlnt::timedelta::daysundefined```
#### ```int xlnt::timedelta::hoursundefined```
#### ```int xlnt::timedelta::minutesundefined```
#### ```int xlnt::timedelta::secondsundefined```
#### ```int xlnt::timedelta::microsecondsundefined```
#### ```static timedelta xlnt::timedelta::from_number(long double number)```
#### ```xlnt::timedelta::timedelta()```
#### ```xlnt::timedelta::timedelta(int days_, int hours_, int minutes_, int seconds_, int microseconds_)```
#### ```long double xlnt::timedelta::to_number() const```
### variant
#### ```undefinedundefined```
#### ```xlnt::variant::variant()```
#### ```xlnt::variant::variant(const std::string &value)```
#### ```xlnt::variant::variant(const char *value)```
#### ```xlnt::variant::variant(int value)```
#### ```xlnt::variant::variant(bool value)```
#### ```xlnt::variant::variant(const datetime &value)```
#### ```xlnt::variant::variant(const std::initializer_list< int > &value)```
#### ```xlnt::variant::variant(const std::vector< int > &value)```
#### ```xlnt::variant::variant(const std::initializer_list< const char *> &value)```
#### ```xlnt::variant::variant(const std::vector< const char *> &value)```
#### ```xlnt::variant::variant(const std::initializer_list< std::string > &value)```
#### ```xlnt::variant::variant(const std::vector< std::string > &value)```
#### ```xlnt::variant::variant(const std::vector< variant > &value)```
#### ```bool xlnt::variant::is(type t) const```
#### ```T xlnt::variant::get() const```
#### ```type xlnt::variant::value_type() const```
## Workbook Module
### calculation_properties
#### ```std::size_t xlnt::calculation_properties::calc_idundefined```
Uniquely identifies these calculation properties.
#### ```bool xlnt::calculation_properties::concurrent_calcundefined```
If this is true, concurrent calculation is enabled.
### const_worksheet_iterator
#### ```xlnt::const_worksheet_iterator::const_worksheet_iterator(const workbook &wb, std::size_t index)```
#### ```xlnt::const_worksheet_iterator::const_worksheet_iterator(const const_worksheet_iterator &)```
#### ```const_worksheet_iterator& xlnt::const_worksheet_iterator::operator=(const const_worksheet_iterator &)```
#### ```const worksheet xlnt::const_worksheet_iterator::operator*()```
#### ```bool xlnt::const_worksheet_iterator::operator==(const const_worksheet_iterator &comparand) const```
#### ```bool xlnt::const_worksheet_iterator::operator!=(const const_worksheet_iterator &comparand) const```
#### ```const_worksheet_iterator xlnt::const_worksheet_iterator::operator++(int)```
#### ```const_worksheet_iterator& xlnt::const_worksheet_iterator::operator++()```
### document_security
#### ```bool xlnt::document_security::lock_revisionundefined```
#### ```bool xlnt::document_security::lock_structureundefined```
#### ```bool xlnt::document_security::lock_windowsundefined```
#### ```std::string xlnt::document_security::revision_passwordundefined```
#### ```std::string xlnt::document_security::workbook_passwordundefined```
#### ```xlnt::document_security::document_security()```
### external_book
### named_range
#### ```using xlnt::named_range::target =  std::pair<worksheet, range_reference>undefined```
#### ```xlnt::named_range::named_range()```
#### ```xlnt::named_range::named_range(const named_range &other)```
#### ```xlnt::named_range::named_range(const std::string &name, const std::vector< target > &targets)```
#### ```std::string xlnt::named_range::name() const```
#### ```const std::vector<target>& xlnt::named_range::targets() const```
#### ```named_range& xlnt::named_range::operator=(const named_range &other)```
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
#### ```void swap(workbook &left, workbook &right)```
Swaps the data held in workbooks "left" and "right".
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
#### ```void xlnt::workbook::save(const std::string &filename, const std::string &password)```
Serializes the workbook into an XLSX file encrypted with the given password and loads the bytes into a file named filename.
#### ```void xlnt::workbook::save(const xlnt::path &filename) const```
Serializes the workbook into an XLSX file and saves the data into a file named filename.
#### ```void xlnt::workbook::save(const xlnt::path &filename, const std::string &password)```
Serializes the workbook into an XLSX file encrypted with the given password and loads the bytes into a file named filename.
#### ```void xlnt::workbook::save(std::ostream &stream) const```
Serializes the workbook into an XLSX file and saves the data into stream.
#### ```void xlnt::workbook::save(std::ostream &stream, const std::string &password)```
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
#### ```bool xlnt::workbook_view::minimizedundefined```
#### ```bool xlnt::workbook_view::show_horizontal_scrollundefined```
#### ```bool xlnt::workbook_view::show_sheet_tabsundefined```
#### ```bool xlnt::workbook_view::show_vertical_scrollundefined```
#### ```bool xlnt::workbook_view::visibleundefined```
#### ```optional<std::size_t> xlnt::workbook_view::active_tabundefined```
#### ```optional<std::size_t> xlnt::workbook_view::first_sheetundefined```
#### ```optional<std::size_t> xlnt::workbook_view::tab_ratioundefined```
#### ```optional<std::size_t> xlnt::workbook_view::window_widthundefined```
#### ```optional<std::size_t> xlnt::workbook_view::window_heightundefined```
#### ```optional<std::size_t> xlnt::workbook_view::x_windowundefined```
#### ```optional<std::size_t> xlnt::workbook_view::y_windowundefined```
### worksheet_iterator
#### ```xlnt::worksheet_iterator::worksheet_iterator(workbook &wb, std::size_t index)```
#### ```xlnt::worksheet_iterator::worksheet_iterator(const worksheet_iterator &)```
#### ```worksheet_iterator& xlnt::worksheet_iterator::operator=(const worksheet_iterator &)```
#### ```worksheet xlnt::worksheet_iterator::operator*()```
#### ```bool xlnt::worksheet_iterator::operator==(const worksheet_iterator &comparand) const```
#### ```bool xlnt::worksheet_iterator::operator!=(const worksheet_iterator &comparand) const```
#### ```worksheet_iterator xlnt::worksheet_iterator::operator++(int)```
#### ```worksheet_iterator& xlnt::worksheet_iterator::operator++()```
## Worksheet Module
### cell_iterator
#### ```xlnt::cell_iterator::cell_iterator(worksheet ws, const cell_reference &start_cell, const range_reference &limits)```
#### ```xlnt::cell_iterator::cell_iterator(worksheet ws, const cell_reference &start_cell, const range_reference &limits, major_order order)```
#### ```xlnt::cell_iterator::cell_iterator(const cell_iterator &other)```
#### ```cell xlnt::cell_iterator::operator*()```
#### ```cell_iterator& xlnt::cell_iterator::operator=(const cell_iterator &)=default```
#### ```bool xlnt::cell_iterator::operator==(const cell_iterator &other) const```
#### ```bool xlnt::cell_iterator::operator!=(const cell_iterator &other) const```
#### ```cell_iterator& xlnt::cell_iterator::operator--()```
#### ```cell_iterator xlnt::cell_iterator::operator--(int)```
#### ```cell_iterator& xlnt::cell_iterator::operator++()```
#### ```cell_iterator xlnt::cell_iterator::operator++(int)```
### cell_vector
#### ```using xlnt::cell_vector::iterator =  cell_iteratorundefined```
#### ```using xlnt::cell_vector::const_iterator =  const_cell_iteratorundefined```
#### ```using xlnt::cell_vector::reverse_iterator =  std::reverse_iterator<iterator>undefined```
#### ```using xlnt::cell_vector::const_reverse_iterator =  std::reverse_iterator<const_iterator>undefined```
#### ```xlnt::cell_vector::cell_vector(worksheet ws, const range_reference &ref, major_order order=major_order::row)```
#### ```cell xlnt::cell_vector::front()```
#### ```const cell xlnt::cell_vector::front() const```
#### ```cell xlnt::cell_vector::back()```
#### ```const cell xlnt::cell_vector::back() const```
#### ```cell xlnt::cell_vector::operator[](std::size_t column_index)```
#### ```const cell xlnt::cell_vector::operator[](std::size_t column_index) const```
#### ```std::size_t xlnt::cell_vector::length() const```
#### ```iterator xlnt::cell_vector::begin()```
#### ```iterator xlnt::cell_vector::end()```
#### ```const_iterator xlnt::cell_vector::begin() const```
#### ```const_iterator xlnt::cell_vector::cbegin() const```
#### ```const_iterator xlnt::cell_vector::end() const```
#### ```const_iterator xlnt::cell_vector::cend() const```
#### ```reverse_iterator xlnt::cell_vector::rbegin()```
#### ```reverse_iterator xlnt::cell_vector::rend()```
#### ```const_reverse_iterator xlnt::cell_vector::rbegin() const```
#### ```const_reverse_iterator xlnt::cell_vector::rend() const```
#### ```const_reverse_iterator xlnt::cell_vector::crbegin() const```
#### ```const_reverse_iterator xlnt::cell_vector::crend() const```
### column_properties
#### ```optional<double> xlnt::column_properties::widthundefined```
#### ```bool xlnt::column_properties::custom_widthundefined```
#### ```optional<std::size_t> xlnt::column_properties::styleundefined```
#### ```bool xlnt::column_properties::hiddenundefined```
### const_cell_iterator
#### ```xlnt::const_cell_iterator::const_cell_iterator(worksheet ws, const cell_reference &start_cell)```
#### ```xlnt::const_cell_iterator::const_cell_iterator(worksheet ws, const cell_reference &start_cell, major_order order)```
#### ```xlnt::const_cell_iterator::const_cell_iterator(const const_cell_iterator &other)```
#### ```const_cell_iterator& xlnt::const_cell_iterator::operator=(const const_cell_iterator &)=default```
#### ```const cell xlnt::const_cell_iterator::operator*() const```
#### ```bool xlnt::const_cell_iterator::operator==(const const_cell_iterator &other) const```
#### ```bool xlnt::const_cell_iterator::operator!=(const const_cell_iterator &other) const```
#### ```const_cell_iterator& xlnt::const_cell_iterator::operator--()```
#### ```const_cell_iterator xlnt::const_cell_iterator::operator--(int)```
#### ```const_cell_iterator& xlnt::const_cell_iterator::operator++()```
#### ```const_cell_iterator xlnt::const_cell_iterator::operator++(int)```
### const_range_iterator
#### ```xlnt::const_range_iterator::const_range_iterator(const worksheet &ws, const range_reference &start_cell, major_order order=major_order::row)```
#### ```xlnt::const_range_iterator::const_range_iterator(const const_range_iterator &other)```
#### ```const cell_vector xlnt::const_range_iterator::operator*() const```
#### ```const_range_iterator& xlnt::const_range_iterator::operator=(const const_range_iterator &)=default```
#### ```bool xlnt::const_range_iterator::operator==(const const_range_iterator &other) const```
#### ```bool xlnt::const_range_iterator::operator!=(const const_range_iterator &other) const```
#### ```const_range_iterator& xlnt::const_range_iterator::operator--()```
#### ```const_range_iterator xlnt::const_range_iterator::operator--(int)```
#### ```const_range_iterator& xlnt::const_range_iterator::operator++()```
#### ```const_range_iterator xlnt::const_range_iterator::operator++(int)```
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
#### ```double xlnt::page_margins::top() const```
#### ```void xlnt::page_margins::top(double top)```
#### ```double xlnt::page_margins::left() const```
#### ```void xlnt::page_margins::left(double left)```
#### ```double xlnt::page_margins::bottom() const```
#### ```void xlnt::page_margins::bottom(double bottom)```
#### ```double xlnt::page_margins::right() const```
#### ```void xlnt::page_margins::right(double right)```
#### ```double xlnt::page_margins::header() const```
#### ```void xlnt::page_margins::header(double header)```
#### ```double xlnt::page_margins::footer() const```
#### ```void xlnt::page_margins::footer(double footer)```
### page_setup
#### ```xlnt::page_setup::page_setup()```
#### ```xlnt::page_break xlnt::page_setup::page_break() const```
#### ```void xlnt::page_setup::page_break(xlnt::page_break b)```
#### ```xlnt::sheet_state xlnt::page_setup::sheet_state() const```
#### ```void xlnt::page_setup::sheet_state(xlnt::sheet_state sheet_state)```
#### ```xlnt::paper_size xlnt::page_setup::paper_size() const```
#### ```void xlnt::page_setup::paper_size(xlnt::paper_size paper_size)```
#### ```xlnt::orientation xlnt::page_setup::orientation() const```
#### ```void xlnt::page_setup::orientation(xlnt::orientation orientation)```
#### ```bool xlnt::page_setup::fit_to_page() const```
#### ```void xlnt::page_setup::fit_to_page(bool fit_to_page)```
#### ```bool xlnt::page_setup::fit_to_height() const```
#### ```void xlnt::page_setup::fit_to_height(bool fit_to_height)```
#### ```bool xlnt::page_setup::fit_to_width() const```
#### ```void xlnt::page_setup::fit_to_width(bool fit_to_width)```
#### ```void xlnt::page_setup::horizontal_centered(bool horizontal_centered)```
#### ```bool xlnt::page_setup::horizontal_centered() const```
#### ```void xlnt::page_setup::vertical_centered(bool vertical_centered)```
#### ```bool xlnt::page_setup::vertical_centered() const```
#### ```void xlnt::page_setup::scale(double scale)```
#### ```double xlnt::page_setup::scale() const```
### pane
#### ```optional<cell_reference> xlnt::pane::top_left_cellundefined```
#### ```pane_state xlnt::pane::stateundefined```
#### ```pane_corner xlnt::pane::active_paneundefined```
#### ```row_t xlnt::pane::y_splitundefined```
#### ```column_t xlnt::pane::x_splitundefined```
#### ```bool xlnt::pane::operator==(const pane &rhs) const```
### range
#### ```using xlnt::range::iterator =  range_iteratorundefined```
#### ```using xlnt::range::const_iterator =  const_range_iteratorundefined```
#### ```using xlnt::range::reverse_iterator =  std::reverse_iterator<iterator>undefined```
#### ```using xlnt::range::const_reverse_iterator =  std::reverse_iterator<const_iterator>undefined```
#### ```xlnt::range::range(worksheet ws, const range_reference &reference, major_order order=major_order::row, bool skip_null=false)```
#### ```xlnt::range::~range()```
#### ```xlnt::range::range(const range &)=default```
#### ```cell_vector xlnt::range::operator[](std::size_t vector_index)```
#### ```const cell_vector xlnt::range::operator[](std::size_t vector_index) const```
#### ```bool xlnt::range::operator==(const range &comparand) const```
#### ```bool xlnt::range::operator!=(const range &comparand) const```
#### ```cell_vector xlnt::range::vector(std::size_t vector_index)```
#### ```const cell_vector xlnt::range::vector(std::size_t vector_index) const```
#### ```class cell xlnt::range::cell(const cell_reference &ref)```
#### ```const class cell xlnt::range::cell(const cell_reference &ref) const```
#### ```range_reference xlnt::range::reference() const```
#### ```std::size_t xlnt::range::length() const```
#### ```bool xlnt::range::contains(const cell_reference &ref)```
#### ```iterator xlnt::range::begin()```
#### ```iterator xlnt::range::end()```
#### ```const_iterator xlnt::range::begin() const```
#### ```const_iterator xlnt::range::end() const```
#### ```const_iterator xlnt::range::cbegin() const```
#### ```const_iterator xlnt::range::cend() const```
#### ```reverse_iterator xlnt::range::rbegin()```
#### ```reverse_iterator xlnt::range::rend()```
#### ```const_reverse_iterator xlnt::range::rbegin() const```
#### ```const_reverse_iterator xlnt::range::rend() const```
#### ```const_reverse_iterator xlnt::range::crbegin() const```
#### ```const_reverse_iterator xlnt::range::crend() const```
### range_iterator
#### ```xlnt::range_iterator::range_iterator(worksheet &ws, const range_reference &start_cell, const range_reference &limits, major_order order=major_order::row)```
#### ```xlnt::range_iterator::range_iterator(const range_iterator &other)```
#### ```cell_vector xlnt::range_iterator::operator*() const```
#### ```range_iterator& xlnt::range_iterator::operator=(const range_iterator &)=default```
#### ```bool xlnt::range_iterator::operator==(const range_iterator &other) const```
#### ```bool xlnt::range_iterator::operator!=(const range_iterator &other) const```
#### ```range_iterator& xlnt::range_iterator::operator--()```
#### ```range_iterator xlnt::range_iterator::operator--(int)```
#### ```range_iterator& xlnt::range_iterator::operator++()```
#### ```range_iterator xlnt::range_iterator::operator++(int)```
### range_reference
#### ```static range_reference xlnt::range_reference::make_absolute(const range_reference &relative_reference)```
Convert a coordinate to an absolute coordinate string (B12 -> $B$12)
#### ```xlnt::range_reference::range_reference()```
#### ```xlnt::range_reference::range_reference(const std::string &range_string)```
#### ```xlnt::range_reference::range_reference(const char *range_string)```
#### ```xlnt::range_reference::range_reference(const std::pair< cell_reference, cell_reference > &reference_pair)```
#### ```xlnt::range_reference::range_reference(const cell_reference &start, const cell_reference &end)```
#### ```xlnt::range_reference::range_reference(column_t column_index_start, row_t row_index_start, column_t column_index_end, row_t row_index_end)```
#### ```bool xlnt::range_reference::is_single_cell() const```
#### ```std::size_t xlnt::range_reference::width() const```
#### ```std::size_t xlnt::range_reference::height() const```
#### ```cell_reference xlnt::range_reference::top_left() const```
#### ```cell_reference xlnt::range_reference::bottom_right() const```
#### ```cell_reference& xlnt::range_reference::top_left()```
#### ```cell_reference& xlnt::range_reference::bottom_right()```
#### ```range_reference xlnt::range_reference::make_offset(int column_offset, int row_offset) const```
#### ```std::string xlnt::range_reference::to_string() const```
#### ```bool xlnt::range_reference::operator==(const range_reference &comparand) const```
#### ```bool xlnt::range_reference::operator==(const std::string &reference_string) const```
#### ```bool xlnt::range_reference::operator==(const char *reference_string) const```
#### ```bool xlnt::range_reference::operator!=(const range_reference &comparand) const```
#### ```bool xlnt::range_reference::operator!=(const std::string &reference_string) const```
#### ```bool xlnt::range_reference::operator!=(const char *reference_string) const```
#### ```bool operator==(const std::string &reference_string, const range_reference &ref)```
#### ```bool operator==(const char *reference_string, const range_reference &ref)```
#### ```bool operator!=(const std::string &reference_string, const range_reference &ref)```
#### ```bool operator!=(const char *reference_string, const range_reference &ref)```
### row_properties
#### ```optional<double> xlnt::row_properties::heightundefined```
#### ```bool xlnt::row_properties::custom_heightundefined```
#### ```bool xlnt::row_properties::hiddenundefined```
#### ```optional<std::size_t> xlnt::row_properties::styleundefined```
### selection
#### ```bool xlnt::selection::has_active_cell() const```
#### ```cell_reference xlnt::selection::active_cell() const```
#### ```void xlnt::selection::active_cell(const cell_reference &ref)```
#### ```range_reference xlnt::selection::sqref() const```
#### ```pane_corner xlnt::selection::pane() const```
#### ```void xlnt::selection::pane(pane_corner corner)```
#### ```bool xlnt::selection::operator==(const selection &rhs) const```
### sheet_protection
#### ```static std::string xlnt::sheet_protection::hash_password(const std::string &password)```
#### ```void xlnt::sheet_protection::password(const std::string &password)```
#### ```std::string xlnt::sheet_protection::hashed_password() const```
### sheet_view
#### ```void xlnt::sheet_view::id(std::size_t new_id)```
#### ```std::size_t xlnt::sheet_view::id() const```
#### ```bool xlnt::sheet_view::has_pane() const```
#### ```struct pane& xlnt::sheet_view::pane()```
#### ```const struct pane& xlnt::sheet_view::pane() const```
#### ```void xlnt::sheet_view::clear_pane()```
#### ```void xlnt::sheet_view::pane(const struct pane &new_pane)```
#### ```bool xlnt::sheet_view::has_selections() const```
#### ```void xlnt::sheet_view::add_selection(const class selection &new_selection)```
#### ```void xlnt::sheet_view::clear_selections()```
#### ```std::vector<xlnt::selection> xlnt::sheet_view::selections() const```
#### ```class xlnt::selection& xlnt::sheet_view::selection(std::size_t index)```
#### ```void xlnt::sheet_view::show_grid_lines(bool show)```
#### ```bool xlnt::sheet_view::show_grid_lines() const```
#### ```void xlnt::sheet_view::default_grid_color(bool is_default)```
#### ```bool xlnt::sheet_view::default_grid_color() const```
#### ```void xlnt::sheet_view::type(sheet_view_type new_type)```
#### ```sheet_view_type xlnt::sheet_view::type() const```
#### ```bool xlnt::sheet_view::operator==(const sheet_view &rhs) const```
### worksheet
#### ```using xlnt::worksheet::iterator =  range_iteratorundefined```
Iterate over a non-const worksheet with an iterator of this type.
#### ```using xlnt::worksheet::const_iterator =  const_range_iteratorundefined```
Iterate over a non-const worksheet with an iterator of this type.
#### ```using xlnt::worksheet::reverse_iterator =  std::reverse_iterator<iterator>undefined```
Iterate in reverse over a non-const worksheet with an iterator of this type.
#### ```using xlnt::worksheet::const_reverse_iterator =  std::reverse_iterator<const_iterator>undefined```
Iterate in reverse order over a const worksheet with an iterator of this type.
#### ```class range const cell_reference& xlnt::worksheet::bottom_rightundefined```
#### ```const class range const cell_reference& bottom_right xlnt::worksheet::constundefined```
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
#### ```class cell xlnt::worksheet::cell(column_t column, row_t row)```
#### ```const class cell xlnt::worksheet::cell(column_t column, row_t row) const```
#### ```class cell xlnt::worksheet::cell(const cell_reference &reference)```
#### ```const class cell xlnt::worksheet::cell(const cell_reference &reference) const```
#### ```bool xlnt::worksheet::has_cell(const cell_reference &reference) const```
#### ```class range xlnt::worksheet::range(const std::string &reference_string)```
#### ```class range xlnt::worksheet::range(const range_reference &reference)```
#### ```const class range xlnt::worksheet::range(const std::string &reference_string) const```
#### ```const class range xlnt::worksheet::range(const range_reference &reference) const```
#### ```class range xlnt::worksheet::rows() const```
#### ```class range xlnt::worksheet::rows(const std::string &range_string) const```
#### ```class range xlnt::worksheet::rows(int row_offset, int column_offset) const```
#### ```class range xlnt::worksheet::rows(const std::string &range_string, int row_offset, int column_offset) const```
#### ```class range xlnt::worksheet::columns() const```
#### ```xlnt::column_properties& xlnt::worksheet::column_properties(column_t column)```
#### ```const xlnt::column_properties& xlnt::worksheet::column_properties(column_t column) const```
#### ```bool xlnt::worksheet::has_column_properties(column_t column) const```
#### ```void xlnt::worksheet::add_column_properties(column_t column, const class column_properties &props)```
#### ```double xlnt::worksheet::column_width(column_t column) const```
Calculate the width of the given column. This will be the default column width if a custom width is not set on this column's column_properties.
#### ```xlnt::row_properties& xlnt::worksheet::row_properties(row_t row)```
#### ```const xlnt::row_properties& xlnt::worksheet::row_properties(row_t row) const```
#### ```bool xlnt::worksheet::has_row_properties(row_t row) const```
#### ```void xlnt::worksheet::add_row_properties(row_t row, const class row_properties &props)```
#### ```double xlnt::worksheet::row_height(row_t row) const```
Calculate the height of the given row. This will be the default row height if a custom height is not set on this row's row_properties.
#### ```cell_reference xlnt::worksheet::point_pos(int left, int top) const```
#### ```cell_reference xlnt::worksheet::point_pos(const std::pair< int, int > &point) const```
#### ```std::string xlnt::worksheet::unique_sheet_name(const std::string &value) const```
#### ```void xlnt::worksheet::create_named_range(const std::string &name, const std::string &reference_string)```
#### ```void xlnt::worksheet::create_named_range(const std::string &name, const range_reference &reference)```
#### ```bool xlnt::worksheet::has_named_range(const std::string &name)```
#### ```class range xlnt::worksheet::named_range(const std::string &name)```
#### ```void xlnt::worksheet::remove_named_range(const std::string &name)```
#### ```row_t xlnt::worksheet::lowest_row() const```
#### ```row_t xlnt::worksheet::highest_row() const```
#### ```row_t xlnt::worksheet::next_row() const```
#### ```column_t xlnt::worksheet::lowest_column() const```
#### ```column_t xlnt::worksheet::highest_column() const```
#### ```range_reference xlnt::worksheet::calculate_dimension() const```
#### ```void xlnt::worksheet::merge_cells(const std::string &reference_string)```
#### ```void xlnt::worksheet::merge_cells(const range_reference &reference)```
#### ```void xlnt::worksheet::merge_cells(column_t start_column, row_t start_row, column_t end_column, row_t end_row)```
#### ```void xlnt::worksheet::unmerge_cells(const std::string &reference_string)```
#### ```void xlnt::worksheet::unmerge_cells(const range_reference &reference)```
#### ```void xlnt::worksheet::unmerge_cells(column_t start_column, row_t start_row, column_t end_column, row_t end_row)```
#### ```std::vector<range_reference> xlnt::worksheet::merged_ranges() const```
#### ```void xlnt::worksheet::append()```
#### ```void xlnt::worksheet::append(const std::vector< std::string > &cells)```
#### ```void xlnt::worksheet::append(const std::vector< int > &cells)```
#### ```void xlnt::worksheet::append(const std::unordered_map< std::string, std::string > &cells)```
#### ```void xlnt::worksheet::append(const std::unordered_map< int, std::string > &cells)```
#### ```void xlnt::worksheet::append(const std::vector< int >::const_iterator begin, const std::vector< int >::const_iterator end)```
#### ```bool xlnt::worksheet::operator==(const worksheet &other) const```
#### ```bool xlnt::worksheet::operator!=(const worksheet &other) const```
#### ```bool xlnt::worksheet::operator==(std::nullptr_t) const```
#### ```bool xlnt::worksheet::operator!=(std::nullptr_t) const```
#### ```void xlnt::worksheet::operator=(const worksheet &other)```
#### ```class cell xlnt::worksheet::operator[](const cell_reference &reference)```
#### ```const class cell xlnt::worksheet::operator[](const cell_reference &reference) const```
#### ```class range xlnt::worksheet::operator[](const range_reference &reference)```
#### ```const class range xlnt::worksheet::operator[](const range_reference &reference) const```
#### ```class range xlnt::worksheet::operator[](const std::string &range_string)```
#### ```const class range xlnt::worksheet::operator[](const std::string &range_string) const```
#### ```class range xlnt::worksheet::operator()(const cell_reference &top_left```
#### ```const class range xlnt::worksheet::operator()(const cell_reference &top_left```
#### ```bool xlnt::worksheet::compare(const worksheet &other, bool reference) const```
#### ```bool xlnt::worksheet::has_page_setup() const```
#### ```xlnt::page_setup xlnt::worksheet::page_setup() const```
#### ```void xlnt::worksheet::page_setup(const struct page_setup &setup)```
#### ```bool xlnt::worksheet::has_page_margins() const```
#### ```xlnt::page_margins xlnt::worksheet::page_margins() const```
#### ```void xlnt::worksheet::page_margins(const class page_margins &margins)```
#### ```range_reference xlnt::worksheet::auto_filter() const```
#### ```void xlnt::worksheet::auto_filter(const std::string &range_string)```
#### ```void xlnt::worksheet::auto_filter(const xlnt::range &range)```
#### ```void xlnt::worksheet::auto_filter(const range_reference &reference)```
#### ```void xlnt::worksheet::clear_auto_filter()```
#### ```bool xlnt::worksheet::has_auto_filter() const```
#### ```void xlnt::worksheet::reserve(std::size_t n)```
#### ```bool xlnt::worksheet::has_header_footer() const```
#### ```class header_footer xlnt::worksheet::header_footer() const```
#### ```void xlnt::worksheet::header_footer(const class header_footer &new_header_footer)```
#### ```void xlnt::worksheet::parent(class workbook &wb)```
#### ```std::vector<std::string> xlnt::worksheet::formula_attributes() const```
#### ```xlnt::sheet_state xlnt::worksheet::sheet_state() const```
#### ```void xlnt::worksheet::sheet_state(xlnt::sheet_state state)```
#### ```iterator xlnt::worksheet::begin()```
#### ```iterator xlnt::worksheet::end()```
#### ```const_iterator xlnt::worksheet::begin() const```
#### ```const_iterator xlnt::worksheet::end() const```
#### ```const_iterator xlnt::worksheet::cbegin() const```
#### ```const_iterator xlnt::worksheet::cend() const```
#### ```class range xlnt::worksheet::iter_cells(bool skip_null)```
#### ```void xlnt::worksheet::print_title_rows(row_t first_row, row_t last_row)```
#### ```void xlnt::worksheet::print_title_rows(row_t last_row)```
#### ```void xlnt::worksheet::print_title_cols(column_t first_column, column_t last_column)```
#### ```void xlnt::worksheet::print_title_cols(column_t last_column)```
#### ```std::string xlnt::worksheet::print_titles() const```
#### ```void xlnt::worksheet::print_area(const std::string &print_area)```
#### ```range_reference xlnt::worksheet::print_area() const```
#### ```bool xlnt::worksheet::has_view() const```
#### ```sheet_view xlnt::worksheet::view(std::size_t index=0) const```
#### ```void xlnt::worksheet::add_view(const sheet_view &new_view)```
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
### worksheet_properties
