# cell
## ```using xlnt::cell::type =  cell_typeundefined```
Alias xlnt::cell_type to xlnt::cell::type since it looks nicer.
## ```friend class detail::xlsx_consumerundefined```
## ```friend class detail::xlsx_producerundefined```
## ```friend struct detail::cell_implundefined```
## ```static const std::unordered_map<std::string, int>& xlnt::cell::error_codes()```
Returns a map of error strings such as #DIV/0! and their associated indices.
## ```xlnt::cell::cell(const cell &)=default```
Default copy constructor.
## ```bool xlnt::cell::has_value() const```
Returns true if value has been set and has not been cleared using cell::clear_value().
## ```T xlnt::cell::value() const```
Returns the value of this cell as an instance of type T. Overloads exist for most C++ fundamental types like bool, int, etc. as well as for std::string and xlnt datetime types: date, time, datetime, and timedelta.
## ```void xlnt::cell::clear_value()```
Makes this cell have a value of type null. All other cell attributes are retained.
## ```void xlnt::cell::value(std::nullptr_t)```
Sets the type of this cell to null.
## ```void xlnt::cell::value(bool boolean_value)```
Sets the value of this cell to the given boolean value.
## ```void xlnt::cell::value(int int_value)```
Sets the value of this cell to the given value.
## ```void xlnt::cell::value(unsigned int int_value)```
Sets the value of this cell to the given value.
## ```void xlnt::cell::value(long long int int_value)```
Sets the value of this cell to the given value.
## ```void xlnt::cell::value(unsigned long long int int_value)```
Sets the value of this cell to the given value.
## ```void xlnt::cell::value(float float_value)```
Sets the value of this cell to the given value.
## ```void xlnt::cell::value(double float_value)```
Sets the value of this cell to the given value.
## ```void xlnt::cell::value(long double float_value)```
Sets the value of this cell to the given value.
## ```void xlnt::cell::value(const date &date_value)```
Sets the value of this cell to the given value.
## ```void xlnt::cell::value(const time &time_value)```
Sets the value of this cell to the given value.
## ```void xlnt::cell::value(const datetime &datetime_value)```
Sets the value of this cell to the given value.
## ```void xlnt::cell::value(const timedelta &timedelta_value)```
Sets the value of this cell to the given value.
## ```void xlnt::cell::value(const std::string &string_value)```
Sets the value of this cell to the given value.
## ```void xlnt::cell::value(const char *string_value)```
Sets the value of this cell to the given value.
## ```void xlnt::cell::value(const rich_text &text_value)```
Sets the value of this cell to the given value.
## ```void xlnt::cell::value(const cell other_cell)```
Sets the value and formatting of this cell to that of other_cell.
## ```void xlnt::cell::value(const std::string &string_value, bool infer_type)```
Analyzes string_value to determine its type, convert it to that type, and set the value of this cell to that converted value.
## ```type xlnt::cell::data_type() const```
Returns the type of this cell.
## ```void xlnt::cell::data_type(type t)```
Sets the type of this cell. This should usually be done indirectly by setting the value of the cell to a value of that type.
## ```bool xlnt::cell::garbage_collectible() const```
There's no reason to keep a cell which has no value and is not a placeholder. Returns true if this cell has no value, style, isn't merged, etc.
## ```bool xlnt::cell::is_date() const```
Returns true iff this cell's number format matches a date format.
## ```cell_reference xlnt::cell::reference() const```
Returns a cell_reference that points to the location of this cell.
## ```column_t xlnt::cell::column() const```
Returns the column of this cell.
## ```row_t xlnt::cell::row() const```
Returns the row of this cell.
## ```std::pair<int, int> xlnt::cell::anchor() const```
Returns the location of this cell as an ordered pair (left, top).
## ```std::string xlnt::cell::hyperlink() const```
Returns the URL of this cell's hyperlink.
## ```void xlnt::cell::hyperlink(const std::string &url)```
Adds a hyperlink to this cell pointing to the URL of the given value.
## ```void xlnt::cell::hyperlink(const std::string &url, const std::string &display)```
Adds a hyperlink to this cell pointing to the URI of the given value and sets the text value of the cell to the given parameter.
## ```void xlnt::cell::hyperlink(xlnt::cell target)```
Adds an internal hyperlink to this cell pointing to the given cell.
## ```bool xlnt::cell::has_hyperlink() const```
Returns true if this cell has a hyperlink set.
## ```class alignment xlnt::cell::computed_alignment() const```
Returns the alignment that should be used when displaying this cell graphically based on the workbook default, the cell-level format, and the named style applied to the cell in that order.
## ```class border xlnt::cell::computed_border() const```
Returns the border that should be used when displaying this cell graphically based on the workbook default, the cell-level format, and the named style applied to the cell in that order.
## ```class fill xlnt::cell::computed_fill() const```
Returns the fill that should be used when displaying this cell graphically based on the workbook default, the cell-level format, and the named style applied to the cell in that order.
## ```class font xlnt::cell::computed_font() const```
Returns the font that should be used when displaying this cell graphically based on the workbook default, the cell-level format, and the named style applied to the cell in that order.
## ```class number_format xlnt::cell::computed_number_format() const```
Returns the number format that should be used when displaying this cell graphically based on the workbook default, the cell-level format, and the named style applied to the cell in that order.
## ```class protection xlnt::cell::computed_protection() const```
Returns the protection that should be used when displaying this cell graphically based on the workbook default, the cell-level format, and the named style applied to the cell in that order.
## ```bool xlnt::cell::has_format() const```
Returns true if this cell has had a format applied to it.
## ```const class format xlnt::cell::format() const```
Returns the format applied to this cell. If this cell has no format, an invalid_attribute exception will be thrown.
## ```void xlnt::cell::format(const class format new_format)```
Applies the cell-level formatting of new_format to this cell.
## ```void xlnt::cell::clear_format()```
Removes the cell-level formatting from this cell. This doesn't affect the style that may also be applied to the cell. Throws an invalid_attribute exception if no format is applied.
## ```class number_format xlnt::cell::number_format() const```
Returns the number format of this cell.
## ```void xlnt::cell::number_format(const class number_format &format)```
Creates a new format in the workbook, sets its number_format to the given format, and applies the format to this cell.
## ```class font xlnt::cell::font() const```
Returns the font applied to the text in this cell.
## ```void xlnt::cell::font(const class font &font_)```
Creates a new format in the workbook, sets its font to the given font, and applies the format to this cell.
## ```class fill xlnt::cell::fill() const```
Returns the fill applied to this cell.
## ```void xlnt::cell::fill(const class fill &fill_)```
Creates a new format in the workbook, sets its fill to the given fill, and applies the format to this cell.
## ```class border xlnt::cell::border() const```
Returns the border of this cell.
## ```void xlnt::cell::border(const class border &border_)```
Creates a new format in the workbook, sets its border to the given border, and applies the format to this cell.
## ```class alignment xlnt::cell::alignment() const```
Returns the alignment of the text in this cell.
## ```void xlnt::cell::alignment(const class alignment &alignment_)```
Creates a new format in the workbook, sets its alignment to the given alignment, and applies the format to this cell.
## ```class protection xlnt::cell::protection() const```
Returns the protection of this cell.
## ```void xlnt::cell::protection(const class protection &protection_)```
Creates a new format in the workbook, sets its protection to the given protection, and applies the format to this cell.
## ```bool xlnt::cell::has_style() const```
Returns true if this cell has had a style applied to it.
## ```class style xlnt::cell::style()```
Returns a wrapper pointing to the named style applied to this cell.
## ```const class style xlnt::cell::style() const```
Returns a wrapper pointing to the named style applied to this cell.
## ```void xlnt::cell::style(const class style &new_style)```
Sets the named style applied to this cell to a style named style_name. Equivalent to style(new_style.name()).
## ```void xlnt::cell::style(const std::string &style_name)```
Sets the named style applied to this cell to a style named style_name. If this style has not been previously created in the workbook, a key_not_found exception will be thrown.
## ```void xlnt::cell::clear_style()```
Removes the named style from this cell. An invalid_attribute exception will be thrown if this cell has no style. This will not affect the cell format of the cell.
## ```std::string xlnt::cell::formula() const```
Returns the string representation of the formula applied to this cell.
## ```void xlnt::cell::formula(const std::string &formula)```
Sets the formula of this cell to the given value. This formula string should begin with '='.
## ```void xlnt::cell::clear_formula()```
Removes the formula from this cell. After this is called, has_formula() will return false.
## ```bool xlnt::cell::has_formula() const```
Returns true if this cell has had a formula applied to it.
## ```std::string xlnt::cell::to_string() const```
Returns a string representing the value of this cell. If the data type is not a string, it will be converted according to the number format.
## ```bool xlnt::cell::is_merged() const```
Returns true iff this cell has been merged with one or more surrounding cells.
## ```void xlnt::cell::merged(bool merged)```
Makes this a merged cell iff merged is true. Generally, this shouldn't be called directly. Instead, use worksheet::merge_cells on its parent worksheet.
## ```std::string xlnt::cell::error() const```
Returns the error string that is stored in this cell.
## ```void xlnt::cell::error(const std::string &error)```
Directly assigns the value of this cell to be the given error.
## ```cell xlnt::cell::offset(int column, int row)```
Returns a cell from this cell's parent workbook at a relative offset given by the parameters.
## ```class worksheet xlnt::cell::worksheet()```
Returns the worksheet that owns this cell.
## ```const class worksheet xlnt::cell::worksheet() const```
Returns the worksheet that owns this cell.
## ```class workbook& xlnt::cell::workbook()```
Returns the workbook of the worksheet that owns this cell.
## ```const class workbook& xlnt::cell::workbook() const```
Returns the workbook of the worksheet that owns this cell.
## ```calendar xlnt::cell::base_date() const```
Returns the base date of the parent workbook.
## ```std::string xlnt::cell::check_string(const std::string &to_check)```
Returns to_check after verifying and fixing encoding, size, and illegal characters.
## ```bool xlnt::cell::has_comment()```
Returns true if this cell has a comment applied.
## ```void xlnt::cell::clear_comment()```
Deletes the comment applied to this cell if it exists.
## ```class comment xlnt::cell::comment()```
Gets the comment applied to this cell.
## ```void xlnt::cell::comment(const std::string &text, const std::string &author="Microsoft Office User")```
Creates a new comment with the given text and optional author and applies it to the cell.
## ```void xlnt::cell::comment(const std::string &comment_text, const class font &comment_font, const std::string &author="Microsoft Office User")```
Creates a new comment with the given text, formatting, and optional author and applies it to the cell.
## ```void xlnt::cell::comment(const class comment &new_comment)```
Apply the comment provided as the only argument to the cell.
## ```double xlnt::cell::width() const```
Returns the width of this cell in pixels.
## ```double xlnt::cell::height() const```
Returns the height of this cell in pixels.
## ```cell& xlnt::cell::operator=(const cell &rhs)```
Makes this cell interally point to rhs. The cell data originally pointed to by this cell will be unchanged.
## ```bool xlnt::cell::operator==(const cell &comparand) const```
Returns true if this cell the same cell as comparand (compared by reference).
## ```bool xlnt::cell::operator==(std::nullptr_t) const```
Returns true if this cell is uninitialized.
