# API Reference
## Cell Module
### cell
#### ```using xlnt::cell::type =  cell_type```
Alias xlnt::cell_type to xlnt::cell::type since it looks nicer.
#### ```friend class detail::xlsx_consumer```
#### ```friend class detail::xlsx_producer```
#### ```friend struct detail::cell_impl```
#### ```static const std::unordered_map<std::string, int>& xlnt::cell::error_codes```
Return a map of error strings such as #DIV/0! and their associated indices.
#### ```xlnt::cell::cell```
Default copy constructor.
#### ```bool xlnt::cell::has_value```
Return true if value has been set and has not been cleared using cell::clear_value().
#### ```T xlnt::cell::value```
Return the value of this cell as an instance of type T. Overloads exist for most C++ fundamental types like bool, int, etc. as well as for std::string and xlnt datetime types: date, time, datetime, and timedelta.
#### ```void xlnt::cell::clear_value```
Make this cell have a value of type null. All other cell attributes are retained.
#### ```void xlnt::cell::value```
Set the value of this cell to the given value. Overloads exist for most C++ fundamental types like bool, int, etc. as well as for std::string and xlnt datetime types: date, time, datetime, and timedelta.
#### ```void xlnt::cell::value```
Analyze string_value to determine its type, convert it to that type, and set the value of this cell to that converted value.
#### ```type xlnt::cell::data_type```
Return the type of this cell.
#### ```void xlnt::cell::data_type```
Set the type of this cell.
#### ```bool xlnt::cell::garbage_collectible```
There's no reason to keep a cell which has no value and is not a placeholder. Return true if this cell has no value, style, isn't merged, etc.
#### ```bool xlnt::cell::is_date```
Return true iff this cell's number format matches a date format.
#### ```cell_reference xlnt::cell::reference```
Return a cell_reference that points to the location of this cell.
#### ```column_t xlnt::cell::column```
Return the column of this cell.
#### ```row_t xlnt::cell::row```
Return the row of this cell.
#### ```std::pair<int, int> xlnt::cell::anchor```
Return the location of this cell as an ordered pair.
#### ```std::string xlnt::cell::hyperlink```
Return the URL of this cell's hyperlink.
#### ```void xlnt::cell::hyperlink```
Add a hyperlink to this cell pointing to the URI of the given value.
#### ```bool xlnt::cell::has_hyperlink```
Return true if this cell has a hyperlink set.
#### ```class alignment xlnt::cell::computed_alignment```
Returns the result of computed_format().alignment().
#### ```class border xlnt::cell::computed_border```
Returns the result of computed_format().border().
#### ```class fill xlnt::cell::computed_fill```
Returns the result of computed_format().fill().
#### ```class font xlnt::cell::computed_font```
Returns the result of computed_format().font().
#### ```class number_format xlnt::cell::computed_number_format```
Returns the result of computed_format().number_format().
#### ```class protection xlnt::cell::computed_protection```
Returns the result of computed_format().protection().
#### ```bool xlnt::cell::has_format```
Return true if this cell has had a format applied to it.
#### ```const class format xlnt::cell::format```
Return a reference to the format applied to this cell. If this cell has no format, an invalid_attribute exception will be thrown.
#### ```void xlnt::cell::format```
Applies the cell-level formatting of new_format to this cell.
#### ```void xlnt::cell::clear_format```
Remove the cell-level formatting from this cell. This doesn't affect the style that may also be applied to the cell. Throws an invalid_attribute exception if no format is applied.
#### ```class number_format xlnt::cell::number_format```
Returns the number format of this cell.
#### ```void xlnt::cell::number_format```
Creates a new format in the workbook, sets its number_format to the given format, and applies the format to this cell.
#### ```class font xlnt::cell::font```
Returns the font applied to the text in this cell.
#### ```void xlnt::cell::font```
Creates a new format in the workbook, sets its font to the given font, and applies the format to this cell.
#### ```class fill xlnt::cell::fill```
Returns the fill applied to this cell.
#### ```void xlnt::cell::fill```
Creates a new format in the workbook, sets its fill to the given fill, and applies the format to this cell.
#### ```class border xlnt::cell::border```
Returns the border of this cell.
#### ```void xlnt::cell::border```
Creates a new format in the workbook, sets its border to the given border, and applies the format to this cell.
#### ```class alignment xlnt::cell::alignment```
Returns the alignment of the text in this cell.
#### ```void xlnt::cell::alignment```
Creates a new format in the workbook, sets its alignment to the given alignment, and applies the format to this cell.
#### ```class protection xlnt::cell::protection```
Returns the protection of this cell.
#### ```void xlnt::cell::protection```
Creates a new format in the workbook, sets its protection to the given protection, and applies the format to this cell.
#### ```bool xlnt::cell::has_style```
Returns true if this cell has had a style applied to it.
#### ```const class style xlnt::cell::style```
Returns a wrapper pointing to the named style applied to this cell.
#### ```void xlnt::cell::style```
Equivalent to style(new_style.name())
#### ```void xlnt::cell::style```
Sets the named style applied to this cell to a style named style_name. If this style has not been previously created in the workbook, a key_not_found exception will be thrown.
#### ```void xlnt::cell::clear_style```
Removes the named style from this cell. An invalid_attribute exception will be thrown if this cell has no style. This will not affect the cell format of the cell.
#### ```std::string xlnt::cell::formula```
Returns the string representation of the formula applied to this cell.
#### ```void xlnt::cell::formula```
Sets the formula of this cell to the given value. This formula string should begin with '='.
#### ```void xlnt::cell::clear_formula```
Removes the formula from this cell. After this is called, has_formula() will return false.
#### ```bool xlnt::cell::has_formula```
Returns true if this cell has had a formula applied to it.
#### ```std::string xlnt::cell::to_string```
Returns a string representing the value of this cell. If the data type is not a string, it will be converted according to the number format.
#### ```bool xlnt::cell::is_merged```
Return true iff this cell has been merged with one or more surrounding cells.
#### ```void xlnt::cell::merged```
Make this a merged cell iff merged is true. Generally, this shouldn't be called directly. Instead, use worksheet::merge_cells on its parent worksheet.
#### ```std::string xlnt::cell::error```
Return the error string that is stored in this cell.
#### ```void xlnt::cell::error```
Directly assign the value of this cell to be the given error.
#### ```cell xlnt::cell::offset```
Return a cell from this cell's parent workbook at a relative offset given by the parameters.
#### ```class worksheet xlnt::cell::worksheet```
Return the worksheet that owns this cell.
#### ```const class worksheet xlnt::cell::worksheet```
Return the worksheet that owns this cell.
#### ```class workbook& xlnt::cell::workbook```
Return the workbook of the worksheet that owns this cell.
#### ```const class workbook& xlnt::cell::workbook```
Return the workbook of the worksheet that owns this cell.
#### ```calendar xlnt::cell::base_date```
Shortcut to return the base date of the parent workbook. Equivalent to workbook().properties().excel_base_date
#### ```std::string xlnt::cell::check_string```
Return to_check after checking encoding, size, and illegal characters.
#### ```bool xlnt::cell::has_comment```
Return true if this cell has a comment applied.
#### ```void xlnt::cell::clear_comment```
Delete the comment applied to this cell if it exists.
#### ```class comment xlnt::cell::comment```
Get the comment applied to this cell.
#### ```void xlnt::cell::comment```
Create a new comment with the given text and optional author and apply it to the cell.
#### ```void xlnt::cell::comment```
Create a new comment with the given text, formatting, and optional author and apply it to the cell.
#### ```void xlnt::cell::comment```
Apply the comment provided as the only argument to the cell.
#### ```double xlnt::cell::width```
#### ```double xlnt::cell::height```
#### ```cell& xlnt::cell::operator=```
Make this cell point to rhs. The cell originally pointed to by this cell will be unchanged.
#### ```bool xlnt::cell::operator==```
Return true if this cell the same cell as comparand (compare by reference).
#### ```bool xlnt::cell::operator==```
Return true if this cell is uninitialized.
### cell_reference_hash
#### ```std::size_t xlnt::cell_reference_hash::operator()```
### cell_reference
#### ```static std::pair<std::string, row_t> xlnt::cell_reference::split_reference```
Split a coordinate string like "A1" into an equivalent pair like {"A", 1}.
#### ```static std::pair<std::string, row_t> xlnt::cell_reference::split_reference```
Split a coordinate string like "A1" into an equivalent pair like {"A", 1}. Reference parameters absolute_column and absolute_row will be set to true if column part or row part are prefixed by a dollar-sign indicating they are absolute, otherwise false.
#### ```xlnt::cell_reference::cell_reference```
Default constructor makes a reference to the top-left-most cell, "A1".
#### ```xlnt::cell_reference::cell_reference```
Constructs a cell_reference from a string reprenting a cell coordinate (e.g. $B14).
#### ```xlnt::cell_reference::cell_reference```
Constructs a cell_reference from a string reprenting a cell coordinate (e.g. $B14).
#### ```xlnt::cell_reference::cell_reference```
Constructs a cell_reference from a 1-indexed column index and row index.
#### ```cell_reference& xlnt::cell_reference::make_absolute```
Convert a coordinate to an absolute coordinate string (e.g. B12 -> $B$12) Defaulting to true, absolute_column and absolute_row can optionally control whether the resulting cell_reference has an absolute column (e.g. B12 -> $B12) and absolute row (e.g. B12 -> B$12) respectively.
#### ```bool xlnt::cell_reference::column_absolute```
Return true if the reference refers to an absolute column, otherwise false.
#### ```void xlnt::cell_reference::column_absolute```
Make this reference have an absolute column if absolute_column is true, otherwise not absolute.
#### ```bool xlnt::cell_reference::row_absolute```
Return true if the reference refers to an absolute row, otherwise false.
#### ```void xlnt::cell_reference::row_absolute```
Make this reference have an absolute row if absolute_row is true, otherwise not absolute.
#### ```column_t xlnt::cell_reference::column```
Return a string that identifies the column of this reference (e.g. second column from left is "B")
#### ```void xlnt::cell_reference::column```
Set the column of this reference from a string that identifies a particular column.
#### ```column_t::index_t xlnt::cell_reference::column_index```
Return a 1-indexed numeric index of the column of this reference.
#### ```void xlnt::cell_reference::column_index```
Set the column of this reference from a 1-indexed number that identifies a particular column.
#### ```row_t xlnt::cell_reference::row```
Return a 1-indexed numeric index of the row of this reference.
#### ```void xlnt::cell_reference::row```
Set the row of this reference from a 1-indexed number that identifies a particular row.
#### ```cell_reference xlnt::cell_reference::make_offset```
Return a cell_reference offset from this cell_reference by the number of columns and rows specified by the parameters. A negative value for column_offset or row_offset results in a reference above or left of this cell_reference, respectively.
#### ```std::string xlnt::cell_reference::to_string```
Return a string like "A1" for cell_reference(1, 1).
#### ```range_reference xlnt::cell_reference::to_range```
Return a 1x1 range_reference containing only this cell_reference.
#### ```range_reference xlnt::cell_reference::operator,```
I've always wanted to overload the comma operator. cell_reference("A", 1), cell_reference("B", 1) will return range_reference(cell_reference("A", 1), cell_reference("B", 1))
#### ```bool xlnt::cell_reference::operator==```
Return true if this reference is identical to comparand including in absoluteness of column and row.
#### ```bool xlnt::cell_reference::operator==```
Construct a cell_reference from reference_string and return the result of their comparison.
#### ```bool xlnt::cell_reference::operator==```
Construct a cell_reference from reference_string and return the result of their comparison.
#### ```bool xlnt::cell_reference::operator!=```
Return true if this reference is not identical to comparand including in absoluteness of column and row.
#### ```bool xlnt::cell_reference::operator!=```
Construct a cell_reference from reference_string and return the result of their comparison.
#### ```bool xlnt::cell_reference::operator!=```
Construct a cell_reference from reference_string and return the result of their comparison.
### comment
#### ```xlnt::comment::comment```
Constructs a new blank comment.
#### ```xlnt::comment::comment```
Constructs a new comment with the given text and author.
#### ```xlnt::comment::comment```
Constructs a new comment with the given unformatted text and author.
#### ```rich_text xlnt::comment::text```
Return the text that will be displayed for this comment.
#### ```std::string xlnt::comment::plain_text```
Return the plain text that will be displayed for this comment without formatting information.
#### ```std::string xlnt::comment::author```
Return the author of this comment.
#### ```void xlnt::comment::hide```
Make this comment only visible when the associated cell is hovered.
#### ```void xlnt::comment::show```
Make this comment always visible.
#### ```bool xlnt::comment::visible```
Returns true if this comment is not hidden.
#### ```void xlnt::comment::position```
Set the absolute position of this cell to the given coordinates.
#### ```int xlnt::comment::left```
Returns the distance from the left side of the sheet to the left side of the comment.
#### ```int xlnt::comment::top```
Returns the distance from the top of the sheet to the top of the comment.
#### ```void xlnt::comment::size```
Set the size of the comment.
#### ```int xlnt::comment::width```
Returns the width of this comment.
#### ```int xlnt::comment::height```
Returns the height of this comment.
#### ```bool operator==```
Return true if both comments are equivalent.
### column_t
#### ```using xlnt::column_t::index_t =  std::uint32_t```
#### ```index_t xlnt::column_t::index```
Internal numeric value of this column index.
#### ```static index_t xlnt::column_t::column_index_from_string```
Convert a column letter into a column number (e.g. B -> 2)
#### ```static std::string xlnt::column_t::column_string_from_index```
Convert a column number into a column letter (3 -> 'C')
#### ```xlnt::column_t::column_t```
Default column_t is the first (left-most) column.
#### ```xlnt::column_t::column_t```
Construct a column from a number.
#### ```xlnt::column_t::column_t```
Construct a column from a string.
#### ```xlnt::column_t::column_t```
Construct a column from a string.
#### ```xlnt::column_t::column_t```
Copy constructor
#### ```xlnt::column_t::column_t```
Move constructor
#### ```std::string xlnt::column_t::column_string```
Return a string representation of this column index.
#### ```column_t& xlnt::column_t::operator=```
Set this column to be the same as rhs's and return reference to self.
#### ```column_t& xlnt::column_t::operator=```
Set this column to be equal to rhs and return reference to self.
#### ```column_t& xlnt::column_t::operator=```
Set this column to be equal to rhs and return reference to self.
#### ```bool xlnt::column_t::operator==```
Return true if this column refers to the same column as other.
#### ```bool xlnt::column_t::operator!=```
Return true if this column doesn't refer to the same column as other.
#### ```bool xlnt::column_t::operator==```
Return true if this column refers to the same column as other.
#### ```bool xlnt::column_t::operator==```
Return true if this column refers to the same column as other.
#### ```bool xlnt::column_t::operator==```
Return true if this column refers to the same column as other.
#### ```bool xlnt::column_t::operator==```
Return true if this column refers to the same column as other.
#### ```bool xlnt::column_t::operator!=```
Return true if this column doesn't refer to the same column as other.
#### ```bool xlnt::column_t::operator!=```
Return true if this column doesn't refer to the same column as other.
#### ```bool xlnt::column_t::operator!=```
Return true if this column doesn't refer to the same column as other.
#### ```bool xlnt::column_t::operator!=```
Return true if this column doesn't refer to the same column as other.
#### ```bool xlnt::column_t::operator>```
Return true if other is to the right of this column.
#### ```bool xlnt::column_t::operator>=```
Return true if other is to the right of or equal to this column.
#### ```bool xlnt::column_t::operator<```
Return true if other is to the left of this column.
#### ```bool xlnt::column_t::operator<=```
Return true if other is to the left of or equal to this column.
#### ```bool xlnt::column_t::operator>```
Return true if other is to the right of this column.
#### ```bool xlnt::column_t::operator>=```
Return true if other is to the right of or equal to this column.
#### ```bool xlnt::column_t::operator<```
Return true if other is to the left of this column.
#### ```bool xlnt::column_t::operator<=```
Return true if other is to the left of or equal to this column.
#### ```column_t& xlnt::column_t::operator++```
Pre-increment this column, making it point to the column one to the right.
#### ```column_t& xlnt::column_t::operator--```
Pre-deccrement this column, making it point to the column one to the left.
#### ```column_t xlnt::column_t::operator++```
Post-increment this column, making it point to the column one to the right and returning the old column.
#### ```column_t xlnt::column_t::operator--```
Post-decrement this column, making it point to the column one to the left and returning the old column.
#### ```column_t& xlnt::column_t::operator+=```
Add rhs to this column and return a reference to this column.
#### ```column_t& xlnt::column_t::operator-=```
Subtrac rhs from this column and return a reference to this column.
#### ```column_t& xlnt::column_t::operator*=```
Multiply this column by rhs and return a reference to this column.
#### ```column_t& xlnt::column_t::operator/=```
Divide this column by rhs and return a reference to this column.
#### ```column_t& xlnt::column_t::operator%=```
Mod this column by rhs and return a reference to this column.
#### ```column_t operator+```
Return the result of adding rhs to this column.
#### ```column_t operator-```
Return the result of subtracing lhs by rhs column.
#### ```column_t operator*```
Return the result of multiply lhs by rhs column.
#### ```column_t operator/```
Return the result of divide lhs by rhs column.
#### ```column_t operator%```
Return the result of mod lhs by rhs column.
#### ```bool operator>```
Return true if other is to the right of this column.
#### ```bool operator>=```
Return true if other is to the right of or equal to this column.
#### ```bool operator<```
Return true if other is to the left of this column.
#### ```bool operator<=```
Return true if other is to the left of or equal to this column.
#### ```void swap```
Swap the columns that left and right refer to.
### column_hash
#### ```std::size_t xlnt::column_hash::operator()```
### column_t >
#### ```size_t std::hash< xlnt::column_t >::operator()```
### rich_text
#### ```xlnt::rich_text::rich_text```
#### ```xlnt::rich_text::rich_text```
#### ```xlnt::rich_text::rich_text```
#### ```xlnt::rich_text::rich_text```
#### ```void xlnt::rich_text::clear```
Remove all text runs from this text.
#### ```void xlnt::rich_text::plain_text```
Clear any runs in this text and add a single run with default formatting and the given string as its textual content.
#### ```std::string xlnt::rich_text::plain_text```
Combine the textual content of each text run in order and return the result.
#### ```std::vector<rich_text_run> xlnt::rich_text::runs```
Returns a copy of the individual runs that comprise this text.
#### ```void xlnt::rich_text::runs```
Set the runs of this text all at once.
#### ```void xlnt::rich_text::add_run```
Add a new run to the end of the set of runs.
#### ```bool xlnt::rich_text::operator==```
Returns true if the runs that make up this text are identical to those in rhs.
#### ```bool xlnt::rich_text::operator==```
Returns true if this text has a single unformatted run with text equal to rhs.
## Packaging Module
### manifest
#### ```void xlnt::manifest::clear```
Unregisters all default and override type and all relationships and known parts.
#### ```std::vector<path> xlnt::manifest::parts```
Returns the path to all internal package parts registered as a source or target of a relationship.
#### ```bool xlnt::manifest::has_relationship```
Returns true if the manifest contains a relationship with the given type with part as the source.
#### ```bool xlnt::manifest::has_relationship```
Returns true if the manifest contains a relationship with the given type with part as the source.
#### ```class relationship xlnt::manifest::relationship```
Returns the relationship with "source" as the source and with a type of "type". Throws a key_not_found exception if no such relationship is found.
#### ```class relationship xlnt::manifest::relationship```
Returns the relationship with "source" as the source and with an ID of "rel_id". Throws a key_not_found exception if no such relationship is found.
#### ```std::vector<xlnt::relationship> xlnt::manifest::relationships```
Returns all relationship with "source" as the source.
#### ```std::vector<xlnt::relationship> xlnt::manifest::relationships```
Returns all relationships with "source" as the source and with a type of "type".
#### ```path xlnt::manifest::canonicalize```
Returns the canonical path of the chain of relationships by traversing through rels and forming the absolute combined path.
#### ```std::string xlnt::manifest::register_relationship```
#### ```std::string xlnt::manifest::register_relationship```
#### ```std::unordered_map<std::string, std::string> xlnt::manifest::unregister_relationship```
Delete the relationship with the given id from source part. Returns a mapping of relationship IDs since IDs are shifted down. For example, if there are three relationships for part a.xml like [rId1, rId2, rId3] and rId2 is deleted, the resulting map would look like [rId3->rId2].
#### ```std::string xlnt::manifest::content_type```
Given the path to a part, returns the content type of the part as a string.
#### ```bool xlnt::manifest::has_default_type```
Returns true if a default content type for the extension has been registered.
#### ```std::vector<std::string> xlnt::manifest::extensions_with_default_types```
Returns a vector of all extensions with registered default content types.
#### ```std::string xlnt::manifest::default_type```
Returns the registered default content type for parts of the given extension.
#### ```void xlnt::manifest::register_default_type```
Associates the given extension with the given content type.
#### ```void xlnt::manifest::unregister_default_type```
Unregisters the default content type for the given extension.
#### ```bool xlnt::manifest::has_override_type```
Returns true if a content type overriding the default content type has been registered for the given part.
#### ```std::string xlnt::manifest::override_type```
Returns the override content type registered for the given part. Throws key_not_found exception if no override type was found.
#### ```std::vector<path> xlnt::manifest::parts_with_overriden_types```
Returns the path of every part in this manifest with an overriden content type.
#### ```void xlnt::manifest::register_override_type```
Overrides any default type registered for the part's extension with the given content type.
#### ```void xlnt::manifest::unregister_override_type```
Unregisters the overriding content type of the given part.
### relationship
#### ```xlnt::relationship::relationship```
#### ```xlnt::relationship::relationship```
#### ```std::string xlnt::relationship::id```
Returns a string of the form rId# that identifies the relationship.
#### ```relationship_type xlnt::relationship::type```
Returns the type of this relationship.
#### ```xlnt::target_mode xlnt::relationship::target_mode```
Returns whether the target of the relationship is internal or external to the package.
#### ```uri xlnt::relationship::source```
Returns the URI of the package part this relationship points to.
#### ```uri xlnt::relationship::target```
Returns the URI of the package part this relationship points to.
#### ```bool xlnt::relationship::operator==```
Returns true if and only if rhs is equal to this relationship.
#### ```bool xlnt::relationship::operator!=```
Returns true if and only if rhs is not equal to this relationship.
### uri
#### ```xlnt::uri::uri```
#### ```xlnt::uri::uri```
#### ```xlnt::uri::uri```
#### ```xlnt::uri::uri```
#### ```bool xlnt::uri::is_relative```
#### ```bool xlnt::uri::is_absolute```
#### ```std::string xlnt::uri::scheme```
#### ```std::string xlnt::uri::authority```
#### ```bool xlnt::uri::has_authentication```
#### ```std::string xlnt::uri::authentication```
#### ```std::string xlnt::uri::username```
#### ```std::string xlnt::uri::password```
#### ```std::string xlnt::uri::host```
#### ```bool xlnt::uri::has_port```
#### ```std::size_t xlnt::uri::port```
#### ```class path xlnt::uri::path```
#### ```bool xlnt::uri::has_query```
#### ```std::string xlnt::uri::query```
#### ```bool xlnt::uri::has_fragment```
#### ```std::string xlnt::uri::fragment```
#### ```std::string xlnt::uri::to_string```
#### ```uri xlnt::uri::make_absolute```
#### ```uri xlnt::uri::make_reference```
#### ```bool operator==```
## Styles Module
### alignment
#### ```bool xlnt::alignment::shrink```
#### ```alignment& xlnt::alignment::shrink```
#### ```bool xlnt::alignment::wrap```
#### ```alignment& xlnt::alignment::wrap```
#### ```optional<int> xlnt::alignment::indent```
#### ```alignment& xlnt::alignment::indent```
#### ```optional<int> xlnt::alignment::rotation```
#### ```alignment& xlnt::alignment::rotation```
#### ```optional<horizontal_alignment> xlnt::alignment::horizontal```
#### ```alignment& xlnt::alignment::horizontal```
#### ```optional<vertical_alignment> xlnt::alignment::vertical```
#### ```alignment& xlnt::alignment::vertical```
#### ```bool operator==```
Returns true if left is exactly equal to right.
#### ```bool operator!=```
Returns true if left is not exactly equal to right.
### border
#### ```static const std::vector<border_side>& xlnt::border::all_sides```
#### ```xlnt::border::border```
#### ```optional<border_property> xlnt::border::side```
#### ```border& xlnt::border::side```
#### ```optional<diagonal_direction> xlnt::border::diagonal```
#### ```border& xlnt::border::diagonal```
#### ```bool operator==```
Returns true if left is exactly equal to right.
#### ```bool operator!=```
Returns true if left is not exactly equal to right.
### border_property
#### ```optional<class color> xlnt::border::border_property::color```
#### ```border_property& xlnt::border::border_property::color```
#### ```optional<border_style> xlnt::border::border_property::style```
#### ```border_property& xlnt::border::border_property::style```
#### ```bool operator==```
Returns true if left is exactly equal to right.
#### ```bool operator!=```
Returns true if left is not exactly equal to right.
### indexed_color
#### ```xlnt::indexed_color::indexed_color```
#### ```std::size_t xlnt::indexed_color::index```
#### ```void xlnt::indexed_color::index```
### theme_color
#### ```xlnt::theme_color::theme_color```
#### ```std::size_t xlnt::theme_color::index```
#### ```void xlnt::theme_color::index```
### rgb_color
#### ```xlnt::rgb_color::rgb_color```
#### ```xlnt::rgb_color::rgb_color```
#### ```std::string xlnt::rgb_color::hex_string```
#### ```std::uint8_t xlnt::rgb_color::red```
#### ```std::uint8_t xlnt::rgb_color::green```
#### ```std::uint8_t xlnt::rgb_color::blue```
#### ```std::uint8_t xlnt::rgb_color::alpha```
#### ```std::array<std::uint8_t, 3> xlnt::rgb_color::rgb```
#### ```std::array<std::uint8_t, 4> xlnt::rgb_color::rgba```
### color
#### ```static const color xlnt::color::black```
#### ```static const color xlnt::color::white```
#### ```static const color xlnt::color::red```
#### ```static const color xlnt::color::darkred```
#### ```static const color xlnt::color::blue```
#### ```static const color xlnt::color::darkblue```
#### ```static const color xlnt::color::green```
#### ```static const color xlnt::color::darkgreen```
#### ```static const color xlnt::color::yellow```
#### ```static const color xlnt::color::darkyellow```
#### ```xlnt::color::color```
#### ```xlnt::color::color```
#### ```xlnt::color::color```
#### ```xlnt::color::color```
#### ```color_type xlnt::color::type```
#### ```bool xlnt::color::is_auto```
#### ```void xlnt::color::auto_```
#### ```const rgb_color& xlnt::color::rgb```
#### ```const indexed_color& xlnt::color::indexed```
#### ```const theme_color& xlnt::color::theme```
#### ```double xlnt::color::tint```
#### ```void xlnt::color::tint```
#### ```bool xlnt::color::operator==```
#### ```bool xlnt::color::operator!=```
### pattern_fill
#### ```xlnt::pattern_fill::pattern_fill```
#### ```pattern_fill_type xlnt::pattern_fill::type```
#### ```pattern_fill& xlnt::pattern_fill::type```
#### ```optional<color> xlnt::pattern_fill::foreground```
#### ```pattern_fill& xlnt::pattern_fill::foreground```
#### ```optional<color> xlnt::pattern_fill::background```
#### ```pattern_fill& xlnt::pattern_fill::background```
#### ```bool operator==```
Returns true if left is exactly equal to right.
#### ```bool operator!=```
Returns true if left is not exactly equal to right.
### gradient_fill
#### ```xlnt::gradient_fill::gradient_fill```
#### ```gradient_fill_type xlnt::gradient_fill::type```
#### ```gradient_fill& xlnt::gradient_fill::type```
#### ```gradient_fill& xlnt::gradient_fill::degree```
#### ```double xlnt::gradient_fill::degree```
#### ```double xlnt::gradient_fill::left```
#### ```gradient_fill& xlnt::gradient_fill::left```
#### ```double xlnt::gradient_fill::right```
#### ```gradient_fill& xlnt::gradient_fill::right```
#### ```double xlnt::gradient_fill::top```
#### ```gradient_fill& xlnt::gradient_fill::top```
#### ```double xlnt::gradient_fill::bottom```
#### ```gradient_fill& xlnt::gradient_fill::bottom```
#### ```gradient_fill& xlnt::gradient_fill::add_stop```
#### ```gradient_fill& xlnt::gradient_fill::clear_stops```
#### ```std::unordered_map<double, color> xlnt::gradient_fill::stops```
#### ```bool operator==```
Returns true if left is exactly equal to right.
#### ```bool operator!=```
Returns true if left is not exactly equal to right.
### fill
#### ```static fill xlnt::fill::solid```
Helper method for the most common use case, setting the fill color of a cell to a single solid color. The foreground and background colors of a fill are not the same as the foreground and background colors of a cell. When setting a fill color in Excel, a new fill is created with the given color as the fill's fgColor and index color number 64 as the bgColor. This method creates a fill in the same way.
#### ```xlnt::fill::fill```
Constructs a fill initialized as a none-type pattern fill with no foreground or background colors.
#### ```xlnt::fill::fill```
Constructs a fill initialized as a pattern fill based on the given pattern.
#### ```xlnt::fill::fill```
Constructs a fill initialized as a gradient fill based on the given gradient.
#### ```fill_type xlnt::fill::type```
Returns the fill_type of this fill depending on how it was constructed.
#### ```class gradient_fill xlnt::fill::gradient_fill```
Returns the gradient fill represented by this fill. Throws an invalid_attribute exception if this is not a gradient fill.
#### ```class pattern_fill xlnt::fill::pattern_fill```
Returns the pattern fill represented by this fill. Throws an invalid_attribute exception if this is not a pattern fill.
#### ```bool operator==```
Returns true if left is exactly equal to right.
#### ```bool operator!=```
Returns true if left is not exactly equal to right.
### font
#### ```undefined```
#### ```bool operator==```
Returns true if left is exactly equal to right.
#### ```bool operator!=```
Returns true if left is not exactly equal to right.
#### ```xlnt::font::font```
#### ```font& xlnt::font::bold```
#### ```bool xlnt::font::bold```
#### ```font& xlnt::font::superscript```
#### ```bool xlnt::font::superscript```
#### ```font& xlnt::font::italic```
#### ```bool xlnt::font::italic```
#### ```font& xlnt::font::strikethrough```
#### ```bool xlnt::font::strikethrough```
#### ```font& xlnt::font::underline```
#### ```bool xlnt::font::underlined```
#### ```underline_style xlnt::font::underline```
#### ```bool xlnt::font::has_size```
#### ```font& xlnt::font::size```
#### ```double xlnt::font::size```
#### ```bool xlnt::font::has_name```
#### ```font& xlnt::font::name```
#### ```std::string xlnt::font::name```
#### ```bool xlnt::font::has_color```
#### ```font& xlnt::font::color```
#### ```xlnt::color xlnt::font::color```
#### ```bool xlnt::font::has_family```
#### ```font& xlnt::font::family```
#### ```std::size_t xlnt::font::family```
#### ```bool xlnt::font::has_scheme```
#### ```font& xlnt::font::scheme```
#### ```std::string xlnt::font::scheme```
### format
#### ```friend struct detail::stylesheet```
#### ```std::size_t xlnt::format::id```
#### ```class alignment& xlnt::format::alignment```
#### ```const class alignment& xlnt::format::alignment```
#### ```format xlnt::format::alignment```
#### ```bool xlnt::format::alignment_applied```
#### ```class border& xlnt::format::border```
#### ```const class border& xlnt::format::border```
#### ```format xlnt::format::border```
#### ```bool xlnt::format::border_applied```
#### ```class fill& xlnt::format::fill```
#### ```const class fill& xlnt::format::fill```
#### ```format xlnt::format::fill```
#### ```bool xlnt::format::fill_applied```
#### ```class font& xlnt::format::font```
#### ```const class font& xlnt::format::font```
#### ```format xlnt::format::font```
#### ```bool xlnt::format::font_applied```
#### ```class number_format& xlnt::format::number_format```
#### ```const class number_format& xlnt::format::number_format```
#### ```format xlnt::format::number_format```
#### ```bool xlnt::format::number_format_applied```
#### ```class protection& xlnt::format::protection```
#### ```const class protection& xlnt::format::protection```
#### ```format xlnt::format::protection```
#### ```bool xlnt::format::protection_applied```
#### ```bool xlnt::format::has_style```
#### ```void xlnt::format::clear_style```
#### ```format xlnt::format::style```
#### ```format xlnt::format::style```
#### ```const class style xlnt::format::style```
### number_format
#### ```static const number_format xlnt::number_format::general```
#### ```static const number_format xlnt::number_format::text```
#### ```static const number_format xlnt::number_format::number```
#### ```static const number_format xlnt::number_format::number_00```
#### ```static const number_format xlnt::number_format::number_comma_separated1```
#### ```static const number_format xlnt::number_format::percentage```
#### ```static const number_format xlnt::number_format::percentage_00```
#### ```static const number_format xlnt::number_format::date_yyyymmdd2```
#### ```static const number_format xlnt::number_format::date_yymmdd```
#### ```static const number_format xlnt::number_format::date_ddmmyyyy```
#### ```static const number_format xlnt::number_format::date_dmyslash```
#### ```static const number_format xlnt::number_format::date_dmyminus```
#### ```static const number_format xlnt::number_format::date_dmminus```
#### ```static const number_format xlnt::number_format::date_myminus```
#### ```static const number_format xlnt::number_format::date_xlsx14```
#### ```static const number_format xlnt::number_format::date_xlsx15```
#### ```static const number_format xlnt::number_format::date_xlsx16```
#### ```static const number_format xlnt::number_format::date_xlsx17```
#### ```static const number_format xlnt::number_format::date_xlsx22```
#### ```static const number_format xlnt::number_format::date_datetime```
#### ```static const number_format xlnt::number_format::date_time1```
#### ```static const number_format xlnt::number_format::date_time2```
#### ```static const number_format xlnt::number_format::date_time3```
#### ```static const number_format xlnt::number_format::date_time4```
#### ```static const number_format xlnt::number_format::date_time5```
#### ```static const number_format xlnt::number_format::date_time6```
#### ```static bool xlnt::number_format::is_builtin_format```
#### ```static const number_format& xlnt::number_format::from_builtin_id```
#### ```xlnt::number_format::number_format```
#### ```xlnt::number_format::number_format```
#### ```xlnt::number_format::number_format```
#### ```xlnt::number_format::number_format```
#### ```void xlnt::number_format::format_string```
#### ```void xlnt::number_format::format_string```
#### ```std::string xlnt::number_format::format_string```
#### ```bool xlnt::number_format::has_id```
#### ```void xlnt::number_format::id```
#### ```std::size_t xlnt::number_format::id```
#### ```std::string xlnt::number_format::format```
#### ```std::string xlnt::number_format::format```
#### ```bool xlnt::number_format::is_date_format```
#### ```bool operator==```
Returns true if left is exactly equal to right.
#### ```bool operator!=```
Returns true if left is not exactly equal to right.
### protection
#### ```static protection xlnt::protection::unlocked_and_visible```
#### ```static protection xlnt::protection::locked_and_visible```
#### ```static protection xlnt::protection::unlocked_and_hidden```
#### ```static protection xlnt::protection::locked_and_hidden```
#### ```xlnt::protection::protection```
#### ```bool xlnt::protection::locked```
#### ```protection& xlnt::protection::locked```
#### ```bool xlnt::protection::hidden```
#### ```protection& xlnt::protection::hidden```
#### ```bool operator==```
Returns true if left is exactly equal to right.
#### ```bool operator!=```
Returns true if left is not exactly equal to right.
### style
#### ```friend struct detail::stylesheet```
#### ```friend class detail::xlsx_consumer```
#### ```xlnt::style::style```
Delete zero-argument constructor
#### ```xlnt::style::style```
Default copy constructor
#### ```std::string xlnt::style::name```
Return the name of this style.
#### ```style xlnt::style::name```
#### ```bool xlnt::style::hidden```
#### ```style xlnt::style::hidden```
#### ```optional<bool> xlnt::style::custom```
#### ```style xlnt::style::custom```
#### ```optional<std::size_t> xlnt::style::builtin_id```
#### ```style xlnt::style::builtin_id```
#### ```class alignment& xlnt::style::alignment```
#### ```const class alignment& xlnt::style::alignment```
#### ```style xlnt::style::alignment```
#### ```bool xlnt::style::alignment_applied```
#### ```class border& xlnt::style::border```
#### ```const class border& xlnt::style::border```
#### ```style xlnt::style::border```
#### ```bool xlnt::style::border_applied```
#### ```class fill& xlnt::style::fill```
#### ```const class fill& xlnt::style::fill```
#### ```style xlnt::style::fill```
#### ```bool xlnt::style::fill_applied```
#### ```class font& xlnt::style::font```
#### ```const class font& xlnt::style::font```
#### ```style xlnt::style::font```
#### ```bool xlnt::style::font_applied```
#### ```class number_format& xlnt::style::number_format```
#### ```const class number_format& xlnt::style::number_format```
#### ```style xlnt::style::number_format```
#### ```bool xlnt::style::number_format_applied```
#### ```class protection& xlnt::style::protection```
#### ```const class protection& xlnt::style::protection```
#### ```style xlnt::style::protection```
#### ```bool xlnt::style::protection_applied```
#### ```bool xlnt::style::operator==```
## Utils Module
### date
#### ```int xlnt::date::year```
#### ```int xlnt::date::month```
#### ```int xlnt::date::day```
#### ```static date xlnt::date::today```
Return the current date according to the system time.
#### ```static date xlnt::date::from_number```
Return a date by adding days_since_base_year to base_date. This includes leap years.
#### ```xlnt::date::date```
#### ```int xlnt::date::to_number```
Return the number of days between this date and base_date.
#### ```int xlnt::date::weekday```
#### ```bool xlnt::date::operator==```
Return true if this date is equal to comparand.
### datetime
#### ```int xlnt::datetime::year```
#### ```int xlnt::datetime::month```
#### ```int xlnt::datetime::day```
#### ```int xlnt::datetime::hour```
#### ```int xlnt::datetime::minute```
#### ```int xlnt::datetime::second```
#### ```int xlnt::datetime::microsecond```
#### ```static datetime xlnt::datetime::now```
Return the current date and time according to the system time.
#### ```static datetime xlnt::datetime::today```
Return the current date and time according to the system time. This is equivalent to datetime::now().
#### ```static datetime xlnt::datetime::from_number```
Return a datetime from number by converting the integer part into a date and the fractional part into a time according to date::from_number and time::from_number.
#### ```static datetime xlnt::datetime::from_iso_string```
#### ```xlnt::datetime::datetime```
#### ```xlnt::datetime::datetime```
#### ```std::string xlnt::datetime::to_string```
#### ```std::string xlnt::datetime::to_iso_string```
#### ```long double xlnt::datetime::to_number```
#### ```bool xlnt::datetime::operator==```
#### ```int xlnt::datetime::weekday```
### exception
#### ```xlnt::exception::exception```
#### ```xlnt::exception::exception```
#### ```virtual xlnt::exception::~exception```
#### ```void xlnt::exception::message```
### invalid_parameter
#### ```xlnt::invalid_parameter::invalid_parameter```
#### ```xlnt::invalid_parameter::invalid_parameter```
#### ```virtual xlnt::invalid_parameter::~invalid_parameter```
### invalid_sheet_title
#### ```xlnt::invalid_sheet_title::invalid_sheet_title```
#### ```xlnt::invalid_sheet_title::invalid_sheet_title```
#### ```virtual xlnt::invalid_sheet_title::~invalid_sheet_title```
### missing_number_format
#### ```xlnt::missing_number_format::missing_number_format```
#### ```virtual xlnt::missing_number_format::~missing_number_format```
### invalid_file
#### ```xlnt::invalid_file::invalid_file```
#### ```xlnt::invalid_file::invalid_file```
#### ```virtual xlnt::invalid_file::~invalid_file```
### illegal_character
#### ```xlnt::illegal_character::illegal_character```
#### ```xlnt::illegal_character::illegal_character```
#### ```virtual xlnt::illegal_character::~illegal_character```
### invalid_data_type
#### ```xlnt::invalid_data_type::invalid_data_type```
#### ```xlnt::invalid_data_type::invalid_data_type```
#### ```virtual xlnt::invalid_data_type::~invalid_data_type```
### invalid_column_string_index
#### ```xlnt::invalid_column_string_index::invalid_column_string_index```
#### ```xlnt::invalid_column_string_index::invalid_column_string_index```
#### ```virtual xlnt::invalid_column_string_index::~invalid_column_string_index```
### invalid_cell_reference
#### ```xlnt::invalid_cell_reference::invalid_cell_reference```
#### ```xlnt::invalid_cell_reference::invalid_cell_reference```
#### ```xlnt::invalid_cell_reference::invalid_cell_reference```
#### ```virtual xlnt::invalid_cell_reference::~invalid_cell_reference```
### invalid_attribute
#### ```xlnt::invalid_attribute::invalid_attribute```
#### ```xlnt::invalid_attribute::invalid_attribute```
#### ```virtual xlnt::invalid_attribute::~invalid_attribute```
### key_not_found
#### ```xlnt::key_not_found::key_not_found```
#### ```xlnt::key_not_found::key_not_found```
#### ```virtual xlnt::key_not_found::~key_not_found```
### no_visible_worksheets
#### ```xlnt::no_visible_worksheets::no_visible_worksheets```
#### ```xlnt::no_visible_worksheets::no_visible_worksheets```
#### ```virtual xlnt::no_visible_worksheets::~no_visible_worksheets```
### unhandled_switch_case
#### ```xlnt::unhandled_switch_case::unhandled_switch_case```
#### ```virtual xlnt::unhandled_switch_case::~unhandled_switch_case```
### unsupported
#### ```xlnt::unsupported::unsupported```
#### ```xlnt::unsupported::unsupported```
#### ```virtual xlnt::unsupported::~unsupported```
### optional
#### ```xlnt::optional< T >::optional```
#### ```xlnt::optional< T >::optional```
#### ```xlnt::optional< T >::operator bool```
#### ```bool xlnt::optional< T >::is_set```
#### ```void xlnt::optional< T >::set```
#### ```T& xlnt::optional< T >::get```
#### ```const T& xlnt::optional< T >::get```
#### ```void xlnt::optional< T >::clear```
#### ```optional& xlnt::optional< T >::operator=```
#### ```bool xlnt::optional< T >::operator==```
### path
#### ```static char xlnt::path::system_separator```
The system-specific path separator character (e.g. '/' or '\').
#### ```xlnt::path::path```
Construct an empty path.
#### ```xlnt::path::path```
Counstruct a path from a string representing the path.
#### ```xlnt::path::path```
Construct a path from a string with an explicit directory seprator.
#### ```bool xlnt::path::is_relative```
Return true iff this path doesn't begin with / (or a drive letter on Windows).
#### ```bool xlnt::path::is_absolute```
Return true iff path::is_relative() is false.
#### ```bool xlnt::path::is_root```
Return true iff this path is the root directory.
#### ```path xlnt::path::parent```
Return a new path that points to the directory containing the current path Return the path unchanged if this path is the absolute or relative root.
#### ```std::string xlnt::path::filename```
Return the last component of this path.
#### ```std::string xlnt::path::extension```
Return the part of the path following the last dot in the filename.
#### ```std::pair<std::string, std::string> xlnt::path::split_extension```
Return a pair of strings resulting from splitting the filename on the last dot.
#### ```std::vector<std::string> xlnt::path::split```
Create a string representing this path separated by the provided separator or the system-default separator if not provided.
#### ```std::string xlnt::path::string```
Create a string representing this path separated by the provided separator or the system-default separator if not provided.
#### ```path xlnt::path::resolve```
If this path is relative, append each component of this path to base_path and return the resulting absolute path. Otherwise, the the current path will be returned and base_path will be ignored.
#### ```path xlnt::path::relative_to```
The inverse of path::resolve. Creates a relative path from an absolute path by removing the common root between base_path and this path. If the current path is already relative, return it unchanged.
#### ```bool xlnt::path::exists```
Return true iff the file or directory pointed to by this path exists on the filesystem.
#### ```bool xlnt::path::is_directory```
Return true if the file or directory pointed to by this path is a directory.
#### ```bool xlnt::path::is_file```
Return true if the file or directory pointed to by this path is a regular file.
#### ```std::string xlnt::path::read_contents```
Open the file pointed to by this path and return a string containing the files contents.
#### ```path xlnt::path::append```
Append the provided part to this path and return the result.
#### ```path xlnt::path::append```
Append the provided part to this path and return the result.
#### ```bool operator==```
Returns true if left path is equal to right path.
### path >
#### ```size_t std::hash< xlnt::path >::operator()```
### scoped_enum_hash
#### ```std::size_t xlnt::scoped_enum_hash< Enum >::operator()```
### time
#### ```int xlnt::time::hour```
#### ```int xlnt::time::minute```
#### ```int xlnt::time::second```
#### ```int xlnt::time::microsecond```
#### ```static time xlnt::time::now```
Return the current time according to the system time.
#### ```static time xlnt::time::from_number```
Return a time from a number representing a fraction of a day. The integer part of number will be ignored. 0.5 would return time(12, 0, 0, 0) or noon, halfway through the day.
#### ```xlnt::time::time```
#### ```xlnt::time::time```
#### ```long double xlnt::time::to_number```
#### ```bool xlnt::time::operator==```
### timedelta
#### ```int xlnt::timedelta::days```
#### ```int xlnt::timedelta::hours```
#### ```int xlnt::timedelta::minutes```
#### ```int xlnt::timedelta::seconds```
#### ```int xlnt::timedelta::microseconds```
#### ```static timedelta xlnt::timedelta::from_number```
#### ```xlnt::timedelta::timedelta```
#### ```xlnt::timedelta::timedelta```
#### ```long double xlnt::timedelta::to_number```
### utf8string
#### ```static utf8string xlnt::utf8string::from_utf8```
#### ```static utf8string xlnt::utf8string::from_latin1```
#### ```static utf8string xlnt::utf8string::from_utf16```
#### ```static utf8string xlnt::utf8string::from_utf32```
#### ```bool xlnt::utf8string::is_valid```
### variant
#### ```undefined```
#### ```xlnt::variant::variant```
#### ```xlnt::variant::variant```
#### ```xlnt::variant::variant```
#### ```xlnt::variant::variant```
#### ```xlnt::variant::variant```
#### ```xlnt::variant::variant```
#### ```xlnt::variant::variant```
#### ```xlnt::variant::variant```
#### ```xlnt::variant::variant```
#### ```xlnt::variant::variant```
#### ```xlnt::variant::variant```
#### ```xlnt::variant::variant```
#### ```xlnt::variant::variant```
#### ```bool xlnt::variant::is```
#### ```T xlnt::variant::get```
#### ```type xlnt::variant::value_type```
## Workbook Module
### calculation_properties
#### ```std::size_t xlnt::calculation_properties::calc_id```
#### ```bool xlnt::calculation_properties::concurrent_calc```
### const_worksheet_iterator
#### ```xlnt::const_worksheet_iterator::const_worksheet_iterator```
#### ```xlnt::const_worksheet_iterator::const_worksheet_iterator```
#### ```const_worksheet_iterator& xlnt::const_worksheet_iterator::operator=```
#### ```const worksheet xlnt::const_worksheet_iterator::operator*```
#### ```bool xlnt::const_worksheet_iterator::operator==```
#### ```bool xlnt::const_worksheet_iterator::operator!=```
#### ```const_worksheet_iterator xlnt::const_worksheet_iterator::operator++```
#### ```const_worksheet_iterator& xlnt::const_worksheet_iterator::operator++```
### document_security
#### ```bool xlnt::document_security::lock_revision```
#### ```bool xlnt::document_security::lock_structure```
#### ```bool xlnt::document_security::lock_windows```
#### ```std::string xlnt::document_security::revision_password```
#### ```std::string xlnt::document_security::workbook_password```
#### ```xlnt::document_security::document_security```
### external_book
### named_range
#### ```using xlnt::named_range::target =  std::pair<worksheet, range_reference>```
#### ```xlnt::named_range::named_range```
#### ```xlnt::named_range::named_range```
#### ```xlnt::named_range::named_range```
#### ```std::string xlnt::named_range::name```
#### ```const std::vector<target>& xlnt::named_range::targets```
#### ```named_range& xlnt::named_range::operator=```
### theme
### workbook
#### ```using xlnt::workbook::iterator =  worksheet_iterator```
#### ```using xlnt::workbook::const_iterator =  const_worksheet_iterator```
#### ```using xlnt::workbook::reverse_iterator =  std::reverse_iterator<iterator>```
#### ```using xlnt::workbook::const_reverse_iterator =  std::reverse_iterator<const_iterator>```
#### ```friend class detail::xlsx_consumer```
#### ```friend class detail::xlsx_producer```
#### ```void swap```
Swap the data held in workbooks "left" and "right".
#### ```static workbook xlnt::workbook::empty```
#### ```xlnt::workbook::workbook```
Create a workbook containing a single empty worksheet.
#### ```xlnt::workbook::workbook```
Move construct this workbook from existing workbook "other".
#### ```xlnt::workbook::workbook```
Copy construct this workbook from existing workbook "other".
#### ```xlnt::workbook::~workbook```
Destroy this workbook.
#### ```worksheet xlnt::workbook::create_sheet```
Create a sheet after the last sheet in this workbook and return it.
#### ```worksheet xlnt::workbook::create_sheet```
Create a sheet at the specified index and return it.
#### ```worksheet xlnt::workbook::create_sheet_with_rel```
This should be private...
#### ```void xlnt::workbook::copy_sheet```
Create a new sheet initializing it with all of the data from the provided worksheet.
#### ```void xlnt::workbook::copy_sheet```
Create a new sheet at the specified index initializing it with all of the data from the provided worksheet.
#### ```worksheet xlnt::workbook::active_sheet```
Returns the worksheet that was most recently accessed. This is also the sheet that will be shown when the workbook is opened in the spreadsheet editor program.
#### ```worksheet xlnt::workbook::sheet_by_title```
Return the worksheet with the given name. This may throw an exception if the sheet isn't found. Use workbook::contains(const std::string &) to make sure the sheet exists.
#### ```const worksheet xlnt::workbook::sheet_by_title```
Return the const worksheet with the given name. This may throw an exception if the sheet isn't found. Use workbook::contains(const std::string &) to make sure the sheet exists.
#### ```worksheet xlnt::workbook::sheet_by_index```
Return the worksheet at the given index.
#### ```const worksheet xlnt::workbook::sheet_by_index```
Return the const worksheet at the given index.
#### ```worksheet xlnt::workbook::sheet_by_id```
Return the worksheet with a sheetId of id.
#### ```const worksheet xlnt::workbook::sheet_by_id```
Return the const worksheet with a sheetId of id.
#### ```bool xlnt::workbook::contains```
Return true if this workbook contains a sheet with the given name.
#### ```std::size_t xlnt::workbook::index```
Return the index of the given worksheet. The worksheet must be owned by this workbook.
#### ```void xlnt::workbook::remove_sheet```
Remove the given worksheet from this workbook.
#### ```void xlnt::workbook::clear```
Delete every cell in this worksheet. After this is called, the worksheet will be equivalent to a newly created sheet at the same index and with the same title.
#### ```iterator xlnt::workbook::begin```
Returns an iterator to the first worksheet in this workbook.
#### ```iterator xlnt::workbook::end```
Returns an iterator to the worksheet following the last worksheet of the workbook. This worksheet acts as a placeholder; attempting to access it will cause an exception to be thrown.
#### ```const_iterator xlnt::workbook::begin```
Returns a const iterator to the first worksheet in this workbook.
#### ```const_iterator xlnt::workbook::end```
Returns a const iterator to the worksheet following the last worksheet of the workbook. This worksheet acts as a placeholder; attempting to access it will cause an exception to be thrown.
#### ```const_iterator xlnt::workbook::cbegin```
Returns an iterator to the first worksheet in this workbook.
#### ```const_iterator xlnt::workbook::cend```
Returns a const iterator to the worksheet following the last worksheet of the workbook. This worksheet acts as a placeholder; attempting to access it will cause an exception to be thrown.
#### ```void xlnt::workbook::apply_to_cells```
Apply the function "f" to every non-empty cell in every worksheet in this workbook.
#### ```std::vector<std::string> xlnt::workbook::sheet_titles```
Returns a temporary vector containing the titles of each sheet in the order of the sheets in the workbook.
#### ```std::size_t xlnt::workbook::sheet_count```
Returns the number of sheets in this workbook.
#### ```bool xlnt::workbook::has_core_property```
Returns true if the workbook has the core property with the given name.
#### ```std::vector<xlnt::core_property> xlnt::workbook::core_properties```
#### ```variant xlnt::workbook::core_property```
#### ```void xlnt::workbook::core_property```
#### ```bool xlnt::workbook::has_extended_property```
Returns true if the workbook has the extended property with the given name.
#### ```std::vector<xlnt::extended_property> xlnt::workbook::extended_properties```
#### ```variant xlnt::workbook::extended_property```
#### ```void xlnt::workbook::extended_property```
#### ```bool xlnt::workbook::has_custom_property```
Returns true if the workbook has the custom property with the given name.
#### ```std::vector<std::string> xlnt::workbook::custom_properties```
#### ```variant xlnt::workbook::custom_property```
#### ```void xlnt::workbook::custom_property```
#### ```calendar xlnt::workbook::base_date```
#### ```void xlnt::workbook::base_date```
#### ```bool xlnt::workbook::has_title```
#### ```std::string xlnt::workbook::title```
#### ```void xlnt::workbook::title```
#### ```std::vector<xlnt::named_range> xlnt::workbook::named_ranges```
#### ```void xlnt::workbook::create_named_range```
#### ```void xlnt::workbook::create_named_range```
#### ```bool xlnt::workbook::has_named_range```
#### ```class range xlnt::workbook::named_range```
#### ```void xlnt::workbook::remove_named_range```
#### ```void xlnt::workbook::save```
#### ```void xlnt::workbook::save```
#### ```void xlnt::workbook::save```
#### ```void xlnt::workbook::save```
#### ```void xlnt::workbook::save```
#### ```void xlnt::workbook::save```
#### ```void xlnt::workbook::save```
#### ```void xlnt::workbook::save```
#### ```void xlnt::workbook::load```
#### ```void xlnt::workbook::load```
#### ```void xlnt::workbook::load```
#### ```void xlnt::workbook::load```
#### ```void xlnt::workbook::load```
#### ```void xlnt::workbook::load```
#### ```void xlnt::workbook::load```
#### ```void xlnt::workbook::load```
#### ```bool xlnt::workbook::has_view```
#### ```workbook_view xlnt::workbook::view```
#### ```void xlnt::workbook::view```
#### ```bool xlnt::workbook::has_code_name```
#### ```std::string xlnt::workbook::code_name```
#### ```void xlnt::workbook::code_name```
#### ```bool xlnt::workbook::has_file_version```
#### ```std::string xlnt::workbook::app_name```
#### ```std::size_t xlnt::workbook::last_edited```
#### ```std::size_t xlnt::workbook::lowest_edited```
#### ```std::size_t xlnt::workbook::rup_build```
#### ```bool xlnt::workbook::has_theme```
#### ```const xlnt::theme& xlnt::workbook::theme```
#### ```void xlnt::workbook::theme```
#### ```xlnt::format xlnt::workbook::format```
#### ```const xlnt::format xlnt::workbook::format```
#### ```xlnt::format xlnt::workbook::create_format```
#### ```void xlnt::workbook::clear_formats```
#### ```bool xlnt::workbook::has_style```
#### ```class style xlnt::workbook::style```
#### ```const class style xlnt::workbook::style```
#### ```class style xlnt::workbook::create_style```
#### ```void xlnt::workbook::clear_styles```
#### ```class manifest& xlnt::workbook::manifest```
#### ```const class manifest& xlnt::workbook::manifest```
#### ```void xlnt::workbook::add_shared_string```
#### ```std::vector<rich_text>& xlnt::workbook::shared_strings```
#### ```const std::vector<rich_text>& xlnt::workbook::shared_strings```
#### ```void xlnt::workbook::thumbnail```
#### ```const std::vector<std::uint8_t>& xlnt::workbook::thumbnail```
#### ```bool xlnt::workbook::has_calculation_properties```
#### ```class calculation_properties xlnt::workbook::calculation_properties```
#### ```void xlnt::workbook::calculation_properties```
#### ```workbook& xlnt::workbook::operator=```
Set the contents of this workbook to be equal to those of "other". Other is passed as value to allow for copy-swap idiom.
#### ```worksheet xlnt::workbook::operator[]```
Return the worksheet with a title of "name".
#### ```worksheet xlnt::workbook::operator[]```
Return the worksheet at "index".
#### ```bool xlnt::workbook::operator==```
Return true if this workbook internal implementation points to the same memory as rhs's.
#### ```bool xlnt::workbook::operator!=```
Return true if this workbook internal implementation doesn't point to the same memory as rhs's.
### workbook_view
#### ```bool xlnt::workbook_view::auto_filter_date_grouping```
#### ```bool xlnt::workbook_view::minimized```
#### ```bool xlnt::workbook_view::show_horizontal_scroll```
#### ```bool xlnt::workbook_view::show_sheet_tabs```
#### ```bool xlnt::workbook_view::show_vertical_scroll```
#### ```bool xlnt::workbook_view::visible```
#### ```optional<std::size_t> xlnt::workbook_view::active_tab```
#### ```optional<std::size_t> xlnt::workbook_view::first_sheet```
#### ```optional<std::size_t> xlnt::workbook_view::tab_ratio```
#### ```optional<std::size_t> xlnt::workbook_view::window_width```
#### ```optional<std::size_t> xlnt::workbook_view::window_height```
#### ```optional<std::size_t> xlnt::workbook_view::x_window```
#### ```optional<std::size_t> xlnt::workbook_view::y_window```
### worksheet_iterator
#### ```xlnt::worksheet_iterator::worksheet_iterator```
#### ```xlnt::worksheet_iterator::worksheet_iterator```
#### ```worksheet_iterator& xlnt::worksheet_iterator::operator=```
#### ```worksheet xlnt::worksheet_iterator::operator*```
#### ```bool xlnt::worksheet_iterator::operator==```
#### ```bool xlnt::worksheet_iterator::operator!=```
#### ```worksheet_iterator xlnt::worksheet_iterator::operator++```
#### ```worksheet_iterator& xlnt::worksheet_iterator::operator++```
## Worksheet Module
### cell_iterator
#### ```xlnt::cell_iterator::cell_iterator```
#### ```xlnt::cell_iterator::cell_iterator```
#### ```xlnt::cell_iterator::cell_iterator```
#### ```cell xlnt::cell_iterator::operator*```
#### ```cell_iterator& xlnt::cell_iterator::operator=```
#### ```bool xlnt::cell_iterator::operator==```
#### ```bool xlnt::cell_iterator::operator!=```
#### ```cell_iterator& xlnt::cell_iterator::operator--```
#### ```cell_iterator xlnt::cell_iterator::operator--```
#### ```cell_iterator& xlnt::cell_iterator::operator++```
#### ```cell_iterator xlnt::cell_iterator::operator++```
### cell_vector
#### ```using xlnt::cell_vector::iterator =  cell_iterator```
#### ```using xlnt::cell_vector::const_iterator =  const_cell_iterator```
#### ```using xlnt::cell_vector::reverse_iterator =  std::reverse_iterator<iterator>```
#### ```using xlnt::cell_vector::const_reverse_iterator =  std::reverse_iterator<const_iterator>```
#### ```xlnt::cell_vector::cell_vector```
#### ```cell xlnt::cell_vector::front```
#### ```const cell xlnt::cell_vector::front```
#### ```cell xlnt::cell_vector::back```
#### ```const cell xlnt::cell_vector::back```
#### ```cell xlnt::cell_vector::operator[]```
#### ```const cell xlnt::cell_vector::operator[]```
#### ```std::size_t xlnt::cell_vector::length```
#### ```iterator xlnt::cell_vector::begin```
#### ```iterator xlnt::cell_vector::end```
#### ```const_iterator xlnt::cell_vector::begin```
#### ```const_iterator xlnt::cell_vector::cbegin```
#### ```const_iterator xlnt::cell_vector::end```
#### ```const_iterator xlnt::cell_vector::cend```
#### ```reverse_iterator xlnt::cell_vector::rbegin```
#### ```reverse_iterator xlnt::cell_vector::rend```
#### ```const_reverse_iterator xlnt::cell_vector::rbegin```
#### ```const_reverse_iterator xlnt::cell_vector::rend```
#### ```const_reverse_iterator xlnt::cell_vector::crbegin```
#### ```const_reverse_iterator xlnt::cell_vector::crend```
### column_properties
#### ```optional<double> xlnt::column_properties::width```
#### ```bool xlnt::column_properties::custom_width```
#### ```optional<std::size_t> xlnt::column_properties::style```
#### ```bool xlnt::column_properties::hidden```
### const_cell_iterator
#### ```xlnt::const_cell_iterator::const_cell_iterator```
#### ```xlnt::const_cell_iterator::const_cell_iterator```
#### ```xlnt::const_cell_iterator::const_cell_iterator```
#### ```const_cell_iterator& xlnt::const_cell_iterator::operator=```
#### ```const cell xlnt::const_cell_iterator::operator*```
#### ```bool xlnt::const_cell_iterator::operator==```
#### ```bool xlnt::const_cell_iterator::operator!=```
#### ```const_cell_iterator& xlnt::const_cell_iterator::operator--```
#### ```const_cell_iterator xlnt::const_cell_iterator::operator--```
#### ```const_cell_iterator& xlnt::const_cell_iterator::operator++```
#### ```const_cell_iterator xlnt::const_cell_iterator::operator++```
### const_range_iterator
#### ```xlnt::const_range_iterator::const_range_iterator```
#### ```xlnt::const_range_iterator::const_range_iterator```
#### ```const cell_vector xlnt::const_range_iterator::operator*```
#### ```const_range_iterator& xlnt::const_range_iterator::operator=```
#### ```bool xlnt::const_range_iterator::operator==```
#### ```bool xlnt::const_range_iterator::operator!=```
#### ```const_range_iterator& xlnt::const_range_iterator::operator--```
#### ```const_range_iterator xlnt::const_range_iterator::operator--```
#### ```const_range_iterator& xlnt::const_range_iterator::operator++```
#### ```const_range_iterator xlnt::const_range_iterator::operator++```
### header_footer
#### ```undefined```
#### ```bool xlnt::header_footer::has_header```
True if any text has been added for a header at any location on any page.
#### ```bool xlnt::header_footer::has_footer```
True if any text has been added for a footer at any location on any page.
#### ```bool xlnt::header_footer::align_with_margins```
True if headers and footers should align to the page margins.
#### ```header_footer& xlnt::header_footer::align_with_margins```
Set to true if headers and footers should align to the page margins. Set to false if headers and footers should align to the edge of the page.
#### ```bool xlnt::header_footer::different_odd_even```
True if headers and footers differ based on page number.
#### ```bool xlnt::header_footer::different_first```
True if headers and footers are different on the first page.
#### ```bool xlnt::header_footer::scale_with_doc```
True if headers and footers should scale to match the worksheet.
#### ```header_footer& xlnt::header_footer::scale_with_doc```
Set to true if headers and footers should scale to match the worksheet.
#### ```bool xlnt::header_footer::has_header```
True if any text has been added at the given location on any page.
#### ```void xlnt::header_footer::clear_header```
Remove all headers from all pages.
#### ```void xlnt::header_footer::clear_header```
Remove header at the given location on any page.
#### ```header_footer& xlnt::header_footer::header```
Add a header at the given location with the given text.
#### ```header_footer& xlnt::header_footer::header```
Add a header at the given location with the given text.
#### ```rich_text xlnt::header_footer::header```
Get the text of the header at the given location. If headers are different on odd and even pages, the odd header will be returned.
#### ```bool xlnt::header_footer::has_first_page_header```
True if a header has been set for the first page at any location.
#### ```bool xlnt::header_footer::has_first_page_header```
True if a header has been set for the first page at the given location.
#### ```void xlnt::header_footer::clear_first_page_header```
Remove all headers from the first page.
#### ```void xlnt::header_footer::clear_first_page_header```
Remove header from the first page at the given location.
#### ```header_footer& xlnt::header_footer::first_page_header```
Add a header on the first page at the given location with the given text.
#### ```rich_text xlnt::header_footer::first_page_header```
Get the text of the first page header at the given location. If no first page header has been set, the general header for that location will be returned.
#### ```bool xlnt::header_footer::has_odd_even_header```
True if different headers have been set for odd and even pages.
#### ```bool xlnt::header_footer::has_odd_even_header```
True if different headers have been set for odd and even pages at the given location.
#### ```void xlnt::header_footer::clear_odd_even_header```
Remove odd/even headers at all locations.
#### ```void xlnt::header_footer::clear_odd_even_header```
Remove odd/even headers at the given location.
#### ```header_footer& xlnt::header_footer::odd_even_header```
Add a header for odd pages at the given location with the given text.
#### ```rich_text xlnt::header_footer::odd_header```
Get the text of the odd page header at the given location. If no odd page header has been set, the general header for that location will be returned.
#### ```rich_text xlnt::header_footer::even_header```
Get the text of the even page header at the given location. If no even page header has been set, the general header for that location will be returned.
#### ```bool xlnt::header_footer::has_footer```
True if any text has been added at the given location on any page.
#### ```void xlnt::header_footer::clear_footer```
Remove all footers from all pages.
#### ```void xlnt::header_footer::clear_footer```
Remove footer at the given location on any page.
#### ```header_footer& xlnt::header_footer::footer```
Add a footer at the given location with the given text.
#### ```header_footer& xlnt::header_footer::footer```
Add a footer at the given location with the given text.
#### ```rich_text xlnt::header_footer::footer```
Get the text of the footer at the given location. If footers are different on odd and even pages, the odd footer will be returned.
#### ```bool xlnt::header_footer::has_first_page_footer```
True if a footer has been set for the first page at any location.
#### ```bool xlnt::header_footer::has_first_page_footer```
True if a footer has been set for the first page at the given location.
#### ```void xlnt::header_footer::clear_first_page_footer```
Remove all footers from the first page.
#### ```void xlnt::header_footer::clear_first_page_footer```
Remove footer from the first page at the given location.
#### ```header_footer& xlnt::header_footer::first_page_footer```
Add a footer on the first page at the given location with the given text.
#### ```rich_text xlnt::header_footer::first_page_footer```
Get the text of the first page footer at the given location. If no first page footer has been set, the general footer for that location will be returned.
#### ```bool xlnt::header_footer::has_odd_even_footer```
True if different footers have been set for odd and even pages.
#### ```bool xlnt::header_footer::has_odd_even_footer```
True if different footers have been set for odd and even pages at the given location.
#### ```void xlnt::header_footer::clear_odd_even_footer```
Remove odd/even footers at all locations.
#### ```void xlnt::header_footer::clear_odd_even_footer```
Remove odd/even footers at the given location.
#### ```header_footer& xlnt::header_footer::odd_even_footer```
Add a footer for odd pages at the given location with the given text.
#### ```rich_text xlnt::header_footer::odd_footer```
Get the text of the odd page footer at the given location. If no odd page footer has been set, the general footer for that location will be returned.
#### ```rich_text xlnt::header_footer::even_footer```
Get the text of the even page footer at the given location. If no even page footer has been set, the general footer for that location will be returned.
### page_margins
#### ```xlnt::page_margins::page_margins```
#### ```double xlnt::page_margins::top```
#### ```void xlnt::page_margins::top```
#### ```double xlnt::page_margins::left```
#### ```void xlnt::page_margins::left```
#### ```double xlnt::page_margins::bottom```
#### ```void xlnt::page_margins::bottom```
#### ```double xlnt::page_margins::right```
#### ```void xlnt::page_margins::right```
#### ```double xlnt::page_margins::header```
#### ```void xlnt::page_margins::header```
#### ```double xlnt::page_margins::footer```
#### ```void xlnt::page_margins::footer```
### page_setup
#### ```xlnt::page_setup::page_setup```
#### ```xlnt::page_break xlnt::page_setup::page_break```
#### ```void xlnt::page_setup::page_break```
#### ```xlnt::sheet_state xlnt::page_setup::sheet_state```
#### ```void xlnt::page_setup::sheet_state```
#### ```xlnt::paper_size xlnt::page_setup::paper_size```
#### ```void xlnt::page_setup::paper_size```
#### ```xlnt::orientation xlnt::page_setup::orientation```
#### ```void xlnt::page_setup::orientation```
#### ```bool xlnt::page_setup::fit_to_page```
#### ```void xlnt::page_setup::fit_to_page```
#### ```bool xlnt::page_setup::fit_to_height```
#### ```void xlnt::page_setup::fit_to_height```
#### ```bool xlnt::page_setup::fit_to_width```
#### ```void xlnt::page_setup::fit_to_width```
#### ```void xlnt::page_setup::horizontal_centered```
#### ```bool xlnt::page_setup::horizontal_centered```
#### ```void xlnt::page_setup::vertical_centered```
#### ```bool xlnt::page_setup::vertical_centered```
#### ```void xlnt::page_setup::scale```
#### ```double xlnt::page_setup::scale```
### pane
#### ```optional<cell_reference> xlnt::pane::top_left_cell```
#### ```pane_state xlnt::pane::state```
#### ```pane_corner xlnt::pane::active_pane```
#### ```row_t xlnt::pane::y_split```
#### ```column_t xlnt::pane::x_split```
#### ```bool xlnt::pane::operator==```
### range
#### ```using xlnt::range::iterator =  range_iterator```
#### ```using xlnt::range::const_iterator =  const_range_iterator```
#### ```using xlnt::range::reverse_iterator =  std::reverse_iterator<iterator>```
#### ```using xlnt::range::const_reverse_iterator =  std::reverse_iterator<const_iterator>```
#### ```xlnt::range::range```
#### ```xlnt::range::~range```
#### ```xlnt::range::range```
#### ```cell_vector xlnt::range::operator[]```
#### ```const cell_vector xlnt::range::operator[]```
#### ```bool xlnt::range::operator==```
#### ```bool xlnt::range::operator!=```
#### ```cell_vector xlnt::range::vector```
#### ```const cell_vector xlnt::range::vector```
#### ```class cell xlnt::range::cell```
#### ```const class cell xlnt::range::cell```
#### ```range_reference xlnt::range::reference```
#### ```std::size_t xlnt::range::length```
#### ```bool xlnt::range::contains```
#### ```iterator xlnt::range::begin```
#### ```iterator xlnt::range::end```
#### ```const_iterator xlnt::range::begin```
#### ```const_iterator xlnt::range::end```
#### ```const_iterator xlnt::range::cbegin```
#### ```const_iterator xlnt::range::cend```
#### ```reverse_iterator xlnt::range::rbegin```
#### ```reverse_iterator xlnt::range::rend```
#### ```const_reverse_iterator xlnt::range::rbegin```
#### ```const_reverse_iterator xlnt::range::rend```
#### ```const_reverse_iterator xlnt::range::crbegin```
#### ```const_reverse_iterator xlnt::range::crend```
### range_iterator
#### ```xlnt::range_iterator::range_iterator```
#### ```xlnt::range_iterator::range_iterator```
#### ```cell_vector xlnt::range_iterator::operator*```
#### ```range_iterator& xlnt::range_iterator::operator=```
#### ```bool xlnt::range_iterator::operator==```
#### ```bool xlnt::range_iterator::operator!=```
#### ```range_iterator& xlnt::range_iterator::operator--```
#### ```range_iterator xlnt::range_iterator::operator--```
#### ```range_iterator& xlnt::range_iterator::operator++```
#### ```range_iterator xlnt::range_iterator::operator++```
### range_reference
#### ```static range_reference xlnt::range_reference::make_absolute```
Convert a coordinate to an absolute coordinate string (B12 -> $B$12)
#### ```xlnt::range_reference::range_reference```
#### ```xlnt::range_reference::range_reference```
#### ```xlnt::range_reference::range_reference```
#### ```xlnt::range_reference::range_reference```
#### ```xlnt::range_reference::range_reference```
#### ```xlnt::range_reference::range_reference```
#### ```bool xlnt::range_reference::is_single_cell```
#### ```std::size_t xlnt::range_reference::width```
#### ```std::size_t xlnt::range_reference::height```
#### ```cell_reference xlnt::range_reference::top_left```
#### ```cell_reference xlnt::range_reference::bottom_right```
#### ```cell_reference& xlnt::range_reference::top_left```
#### ```cell_reference& xlnt::range_reference::bottom_right```
#### ```range_reference xlnt::range_reference::make_offset```
#### ```std::string xlnt::range_reference::to_string```
#### ```bool xlnt::range_reference::operator==```
#### ```bool xlnt::range_reference::operator==```
#### ```bool xlnt::range_reference::operator==```
#### ```bool xlnt::range_reference::operator!=```
#### ```bool xlnt::range_reference::operator!=```
#### ```bool xlnt::range_reference::operator!=```
#### ```bool operator==```
#### ```bool operator==```
#### ```bool operator!=```
#### ```bool operator!=```
### row_properties
#### ```optional<double> xlnt::row_properties::height```
#### ```bool xlnt::row_properties::custom_height```
#### ```bool xlnt::row_properties::hidden```
#### ```optional<std::size_t> xlnt::row_properties::style```
### selection
#### ```bool xlnt::selection::has_active_cell```
#### ```cell_reference xlnt::selection::active_cell```
#### ```void xlnt::selection::active_cell```
#### ```range_reference xlnt::selection::sqref```
#### ```pane_corner xlnt::selection::pane```
#### ```void xlnt::selection::pane```
#### ```bool xlnt::selection::operator==```
### sheet_protection
#### ```static std::string xlnt::sheet_protection::hash_password```
#### ```void xlnt::sheet_protection::password```
#### ```std::string xlnt::sheet_protection::hashed_password```
### sheet_view
#### ```void xlnt::sheet_view::id```
#### ```std::size_t xlnt::sheet_view::id```
#### ```bool xlnt::sheet_view::has_pane```
#### ```struct pane& xlnt::sheet_view::pane```
#### ```const struct pane& xlnt::sheet_view::pane```
#### ```void xlnt::sheet_view::clear_pane```
#### ```void xlnt::sheet_view::pane```
#### ```bool xlnt::sheet_view::has_selections```
#### ```void xlnt::sheet_view::add_selection```
#### ```void xlnt::sheet_view::clear_selections```
#### ```std::vector<xlnt::selection> xlnt::sheet_view::selections```
#### ```class xlnt::selection& xlnt::sheet_view::selection```
#### ```void xlnt::sheet_view::show_grid_lines```
#### ```bool xlnt::sheet_view::show_grid_lines```
#### ```void xlnt::sheet_view::default_grid_color```
#### ```bool xlnt::sheet_view::default_grid_color```
#### ```void xlnt::sheet_view::type```
#### ```sheet_view_type xlnt::sheet_view::type```
#### ```bool xlnt::sheet_view::operator==```
### worksheet
#### ```using xlnt::worksheet::iterator =  range_iterator```
Iterate over a non-const worksheet with an iterator of this type.
#### ```using xlnt::worksheet::const_iterator =  const_range_iterator```
Iterate over a non-const worksheet with an iterator of this type.
#### ```using xlnt::worksheet::reverse_iterator =  std::reverse_iterator<iterator>```
Iterate in reverse over a non-const worksheet with an iterator of this type.
#### ```using xlnt::worksheet::const_reverse_iterator =  std::reverse_iterator<const_iterator>```
Iterate in reverse order over a const worksheet with an iterator of this type.
#### ```class range const cell_reference& xlnt::worksheet::bottom_right```
#### ```const class range const cell_reference& bottom_right xlnt::worksheet::const```
#### ```friend class detail::xlsx_consumer```
#### ```friend class detail::xlsx_producer```
#### ```xlnt::worksheet::worksheet```
Construct a null worksheet. No methods should be called on such a worksheet.
#### ```xlnt::worksheet::worksheet```
Copy constructor. This worksheet will point to the same memory as rhs's worksheet.
#### ```class workbook& xlnt::worksheet::workbook```
Returns a reference to the workbook this worksheet is owned by.
#### ```const class workbook& xlnt::worksheet::workbook```
Returns a reference to the workbook this worksheet is owned by.
#### ```void xlnt::worksheet::garbage_collect```
Deletes data held in the worksheet that does not affect the internal data or display. For example, unreference styles and empty cells will be removed.
#### ```std::size_t xlnt::worksheet::id```
Returns the unique numeric identifier of this worksheet. This will sometimes but not necessarily be the index of the worksheet in the workbook.
#### ```void xlnt::worksheet::id```
Set the unique numeric identifier. The id defaults to the lowest unused id in the workbook so this should not be called without a good reason.
#### ```std::string xlnt::worksheet::title```
Returns the title of this sheet.
#### ```void xlnt::worksheet::title```
Sets the title of this sheet.
#### ```cell_reference xlnt::worksheet::frozen_panes```
Returns the top left corner of the region above and to the left of which panes are frozen.
#### ```void xlnt::worksheet::freeze_panes```
Freeze panes above and to the left of top_left_cell.
#### ```void xlnt::worksheet::freeze_panes```
Freeze panes above and to the left of top_left_coordinate.
#### ```void xlnt::worksheet::unfreeze_panes```
Remove frozen panes. The data in those panes will be unaffectedthis affects only the view.
#### ```bool xlnt::worksheet::has_frozen_panes```
Returns true if this sheet has a frozen row, frozen column, or both.
#### ```class cell xlnt::worksheet::cell```
#### ```const class cell xlnt::worksheet::cell```
#### ```class cell xlnt::worksheet::cell```
#### ```const class cell xlnt::worksheet::cell```
#### ```bool xlnt::worksheet::has_cell```
#### ```class range xlnt::worksheet::range```
#### ```class range xlnt::worksheet::range```
#### ```const class range xlnt::worksheet::range```
#### ```const class range xlnt::worksheet::range```
#### ```class range xlnt::worksheet::rows```
#### ```class range xlnt::worksheet::rows```
#### ```class range xlnt::worksheet::rows```
#### ```class range xlnt::worksheet::rows```
#### ```class range xlnt::worksheet::columns```
#### ```xlnt::column_properties& xlnt::worksheet::column_properties```
#### ```const xlnt::column_properties& xlnt::worksheet::column_properties```
#### ```bool xlnt::worksheet::has_column_properties```
#### ```void xlnt::worksheet::add_column_properties```
#### ```double xlnt::worksheet::column_width```
Calculate the width of the given column. This will be the default column width if a custom width is not set on this column's column_properties.
#### ```xlnt::row_properties& xlnt::worksheet::row_properties```
#### ```const xlnt::row_properties& xlnt::worksheet::row_properties```
#### ```bool xlnt::worksheet::has_row_properties```
#### ```void xlnt::worksheet::add_row_properties```
#### ```double xlnt::worksheet::row_height```
Calculate the height of the given row. This will be the default row height if a custom height is not set on this row's row_properties.
#### ```cell_reference xlnt::worksheet::point_pos```
#### ```cell_reference xlnt::worksheet::point_pos```
#### ```std::string xlnt::worksheet::unique_sheet_name```
#### ```void xlnt::worksheet::create_named_range```
#### ```void xlnt::worksheet::create_named_range```
#### ```bool xlnt::worksheet::has_named_range```
#### ```class range xlnt::worksheet::named_range```
#### ```void xlnt::worksheet::remove_named_range```
#### ```row_t xlnt::worksheet::lowest_row```
#### ```row_t xlnt::worksheet::highest_row```
#### ```row_t xlnt::worksheet::next_row```
#### ```column_t xlnt::worksheet::lowest_column```
#### ```column_t xlnt::worksheet::highest_column```
#### ```range_reference xlnt::worksheet::calculate_dimension```
#### ```void xlnt::worksheet::merge_cells```
#### ```void xlnt::worksheet::merge_cells```
#### ```void xlnt::worksheet::merge_cells```
#### ```void xlnt::worksheet::unmerge_cells```
#### ```void xlnt::worksheet::unmerge_cells```
#### ```void xlnt::worksheet::unmerge_cells```
#### ```std::vector<range_reference> xlnt::worksheet::merged_ranges```
#### ```void xlnt::worksheet::append```
#### ```void xlnt::worksheet::append```
#### ```void xlnt::worksheet::append```
#### ```void xlnt::worksheet::append```
#### ```void xlnt::worksheet::append```
#### ```void xlnt::worksheet::append```
#### ```bool xlnt::worksheet::operator==```
#### ```bool xlnt::worksheet::operator!=```
#### ```bool xlnt::worksheet::operator==```
#### ```bool xlnt::worksheet::operator!=```
#### ```void xlnt::worksheet::operator=```
#### ```class cell xlnt::worksheet::operator[]```
#### ```const class cell xlnt::worksheet::operator[]```
#### ```class range xlnt::worksheet::operator[]```
#### ```const class range xlnt::worksheet::operator[]```
#### ```class range xlnt::worksheet::operator[]```
#### ```const class range xlnt::worksheet::operator[]```
#### ```class range xlnt::worksheet::operator```
#### ```const class range xlnt::worksheet::operator```
#### ```bool xlnt::worksheet::compare```
#### ```bool xlnt::worksheet::has_page_setup```
#### ```xlnt::page_setup xlnt::worksheet::page_setup```
#### ```void xlnt::worksheet::page_setup```
#### ```bool xlnt::worksheet::has_page_margins```
#### ```xlnt::page_margins xlnt::worksheet::page_margins```
#### ```void xlnt::worksheet::page_margins```
#### ```range_reference xlnt::worksheet::auto_filter```
#### ```void xlnt::worksheet::auto_filter```
#### ```void xlnt::worksheet::auto_filter```
#### ```void xlnt::worksheet::auto_filter```
#### ```void xlnt::worksheet::clear_auto_filter```
#### ```bool xlnt::worksheet::has_auto_filter```
#### ```void xlnt::worksheet::reserve```
#### ```bool xlnt::worksheet::has_header_footer```
#### ```class header_footer xlnt::worksheet::header_footer```
#### ```void xlnt::worksheet::header_footer```
#### ```void xlnt::worksheet::parent```
#### ```std::vector<std::string> xlnt::worksheet::formula_attributes```
#### ```xlnt::sheet_state xlnt::worksheet::sheet_state```
#### ```void xlnt::worksheet::sheet_state```
#### ```iterator xlnt::worksheet::begin```
#### ```iterator xlnt::worksheet::end```
#### ```const_iterator xlnt::worksheet::begin```
#### ```const_iterator xlnt::worksheet::end```
#### ```const_iterator xlnt::worksheet::cbegin```
#### ```const_iterator xlnt::worksheet::cend```
#### ```class range xlnt::worksheet::iter_cells```
#### ```void xlnt::worksheet::print_title_rows```
#### ```void xlnt::worksheet::print_title_rows```
#### ```void xlnt::worksheet::print_title_cols```
#### ```void xlnt::worksheet::print_title_cols```
#### ```std::string xlnt::worksheet::print_titles```
#### ```void xlnt::worksheet::print_area```
#### ```range_reference xlnt::worksheet::print_area```
#### ```bool xlnt::worksheet::has_view```
#### ```sheet_view xlnt::worksheet::view```
#### ```void xlnt::worksheet::add_view```
#### ```void xlnt::worksheet::clear_page_breaks```
Remove all manual column and row page breaks (represented as dashed blue lines in the page view in Excel).
#### ```const std::vector<row_t>& xlnt::worksheet::page_break_rows```
Returns vector where each element represents a row which will break a page below it.
#### ```void xlnt::worksheet::page_break_at_row```
Add a page break at the given row.
#### ```const std::vector<column_t>& xlnt::worksheet::page_break_columns```
Returns vector where each element represents a column which will break a page to the right.
#### ```void xlnt::worksheet::page_break_at_column```
Add a page break at the given column.
### worksheet_properties
