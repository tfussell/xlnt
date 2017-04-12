# cell_reference
## ```static std::pair<std::string, row_t> xlnt::cell_reference::split_reference(const std::string &reference_string)```
Splits a coordinate string like "A1" into an equivalent pair like {"A", 1}.
## ```static std::pair<std::string, row_t> xlnt::cell_reference::split_reference(const std::string &reference_string, bool &absolute_column, bool &absolute_row)```
Splits a coordinate string like "A1" into an equivalent pair like {"A", 1}. Reference parameters absolute_column and absolute_row will be set to true if column part or row part are prefixed by a dollar-sign indicating they are absolute, otherwise false.
## ```xlnt::cell_reference::cell_reference()```
Default constructor makes a reference to the top-left-most cell, "A1".
## ```xlnt::cell_reference::cell_reference(const char *reference_string)```
Constructs a cell_reference from a string reprenting a cell coordinate (e.g. $B14).
## ```xlnt::cell_reference::cell_reference(const std::string &reference_string)```
Constructs a cell_reference from a string reprenting a cell coordinate (e.g. $B14).
## ```xlnt::cell_reference::cell_reference(column_t column, row_t row)```
Constructs a cell_reference from a 1-indexed column index and row index.
## ```cell_reference& xlnt::cell_reference::make_absolute(bool absolute_column=true, bool absolute_row=true)```
Converts a coordinate to an absolute coordinate string (e.g. B12 -> $B$12) Defaulting to true, absolute_column and absolute_row can optionally control whether the resulting cell_reference has an absolute column (e.g. B12 -> $B12) and absolute row (e.g. B12 -> B$12) respectively.
## ```bool xlnt::cell_reference::column_absolute() const```
Returns true if the reference refers to an absolute column, otherwise false.
## ```void xlnt::cell_reference::column_absolute(bool absolute_column)```
Makes this reference have an absolute column if absolute_column is true, otherwise not absolute.
## ```bool xlnt::cell_reference::row_absolute() const```
Returns true if the reference refers to an absolute row, otherwise false.
## ```void xlnt::cell_reference::row_absolute(bool absolute_row)```
Makes this reference have an absolute row if absolute_row is true, otherwise not absolute.
## ```column_t xlnt::cell_reference::column() const```
Returns a string that identifies the column of this reference (e.g. second column from left is "B")
## ```void xlnt::cell_reference::column(const std::string &column_string)```
Sets the column of this reference from a string that identifies a particular column.
## ```column_t::index_t xlnt::cell_reference::column_index() const```
Returns a 1-indexed numeric index of the column of this reference.
## ```void xlnt::cell_reference::column_index(column_t column)```
Sets the column of this reference from a 1-indexed number that identifies a particular column.
## ```row_t xlnt::cell_reference::row() const```
Returns a 1-indexed numeric index of the row of this reference.
## ```void xlnt::cell_reference::row(row_t row)```
Sets the row of this reference from a 1-indexed number that identifies a particular row.
## ```cell_reference xlnt::cell_reference::make_offset(int column_offset, int row_offset) const```
Returns a cell_reference offset from this cell_reference by the number of columns and rows specified by the parameters. A negative value for column_offset or row_offset results in a reference above or left of this cell_reference, respectively.
## ```std::string xlnt::cell_reference::to_string() const```
Returns a string like "A1" for cell_reference(1, 1).
## ```range_reference xlnt::cell_reference::to_range() const```
Returns a 1x1 range_reference containing only this cell_reference.
## ```range_reference xlnt::cell_reference::operator,(const cell_reference &other) const```
I've always wanted to overload the comma operator. cell_reference("A", 1), cell_reference("B", 1) will return range_reference(cell_reference("A", 1), cell_reference("B", 1))
## ```bool xlnt::cell_reference::operator==(const cell_reference &comparand) const```
Returns true if this reference is identical to comparand including in absoluteness of column and row.
## ```bool xlnt::cell_reference::operator==(const std::string &reference_string) const```
Constructs a cell_reference from reference_string and return the result of their comparison.
## ```bool xlnt::cell_reference::operator==(const char *reference_string) const```
Constructs a cell_reference from reference_string and return the result of their comparison.
## ```bool xlnt::cell_reference::operator!=(const cell_reference &comparand) const```
Returns true if this reference is not identical to comparand including in absoluteness of column and row.
## ```bool xlnt::cell_reference::operator!=(const std::string &reference_string) const```
Constructs a cell_reference from reference_string and return the result of their comparison.
## ```bool xlnt::cell_reference::operator!=(const char *reference_string) const```
Constructs a cell_reference from reference_string and return the result of their comparison.
