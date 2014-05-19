#pragma once

#include <list>
#include <ostream>
#include <string>
#include <unordered_map>
#include <vector>

#include "zip.h"
#include "unzip.h"

struct tm;

namespace xlnt {

class cell;
class comment;
class drawing;
class named_range;
class relationship;
class style;
class workbook;
class worksheet;

struct cell_struct;
struct drawing_struct;
struct named_range_struct;
struct worksheet_struct;

const int MIN_ROW = 0;
const int MIN_COLUMN = 0;
const int MAX_COLUMN = 16384;
const int MAX_ROW = 1048576;

// constants
const std::string PACKAGE_PROPS = "docProps";
const std::string PACKAGE_XL = "xl";
const std::string PACKAGE_RELS = "_rels";
const std::string PACKAGE_THEME = PACKAGE_XL + "/" + "theme";
const std::string PACKAGE_WORKSHEETS = PACKAGE_XL + "/" + "worksheets";
const std::string PACKAGE_DRAWINGS = PACKAGE_XL + "/" + "drawings";
const std::string PACKAGE_CHARTS = PACKAGE_XL + "/" + "charts";

const std::string ARC_CONTENT_TYPES = "[Content_Types].xml";
const std::string ARC_ROOT_RELS = PACKAGE_RELS + "/.rels";
const std::string ARC_WORKBOOK_RELS = PACKAGE_XL + "/" + PACKAGE_RELS + "/workbook.xml.rels";
const std::string ARC_CORE = PACKAGE_PROPS + "/core.xml";
const std::string ARC_APP = PACKAGE_PROPS + "/app.xml";
const std::string ARC_WORKBOOK = PACKAGE_XL + "/workbook.xml";
const std::string ARC_STYLE = PACKAGE_XL + "/styles.xml";
const std::string ARC_THEME = PACKAGE_THEME + "/theme1.xml";
const std::string ARC_SHARED_STRINGS = PACKAGE_XL + "/sharedStrings.xml";

const std::unordered_map<std::string, std::string> NAMESPACES = {
    {"cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"},
    {"dc", "http://purl.org/dc/elements/1.1/"},
    {"dcterms", "http://purl.org/dc/terms/"},
    {"dcmitype", "http://purl.org/dc/dcmitype/"},
    {"xsi", "http://www.w3.org/2001/XMLSchema-instance"},
    {"vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"},
    {"xml", "http://www.w3.org/XML/1998/namespace"}
};

/// <summary>
/// Specifies the compression level for content that is stored in a part.
/// </summary>
enum class compression_option
{
    /// <summary>
    /// Compression is optimized for performance.
    /// </summary>
    Fast,
    /// <summary>
    /// Compression is optimized for size.
    /// </summary>
    Maximum,
    /// <summary>
    /// Compression is optimized for a balance between size and performance.
    /// </summary>
    Normal,
    /// <summary>
    /// Compression is turned off.
    /// </summary>
    NotCompressed,
    /// <summary>
    /// Compression is optimized for high performance.
    /// </summary>
    SuperFast
};

/// <summary>
/// Defines constants for read, write, or read / write access to a file.
/// </summary>
enum class file_access
{
    /// <summary>
    /// Read access to the file. Data can be read from the file. Combine with Write for read/write access.
    /// </summary>
    read = 0x01,
    /// <summary>
    /// Read and write access to the file. Data can be written to and read from the file.
    /// </summary>
    read_write = 0x02,
    /// <summary>
    /// Write access to the file. Data can be written to the file. Combine with Read for read/write access.
    /// </summary>
    write = 0x04
};

/// <summary>
/// Specifies how the operating system should open a file.
/// </summary>
enum class file_mode
{
    /// <summary>
    /// Opens the file if it exists and seeks to the end of the file, or creates a new file.This requires FileIOPermissionAccess.Append permission.file_mode.Append can be used only in conjunction with file_access.Write.Trying to seek to a position before the end of the file throws an IOException exception, and any attempt to read fails and throws a NotSupportedException exception.
    /// </summary>
    append,
    /// <summary>
    /// Specifies that the operating system should create a new file. If the file already exists, it will be overwritten. This requires FileIOPermissionAccess.Write permission. file_mode.Create is equivalent to requesting that if the file does not exist, use CreateNew; otherwise, use Truncate. If the file already exists but is a hidden file, an UnauthorizedAccessException exception is thrown.
    /// </summary>
    create,
    /// <summary>
    /// Specifies that the operating system should create a new file. This requires FileIOPermissionAccess.Write permission. If the file already exists, an IOException exception is thrown.
    /// </summary>
    create_new,
    /// <summary>
    /// Specifies that the operating system should open an existing file. The ability to open the file is dependent on the value specified by the file_access enumeration. A System.IO.FileNotFoundException exception is thrown if the file does not exist.
    /// </summary>
    open,
    /// <summary>
    /// Specifies that the operating system should open a file if it exists; otherwise, a new file should be created. If the file is opened with file_access.Read, FileIOPermissionAccess.Read permission is required. If the file access is file_access.Write, FileIOPermissionAccess.Write permission is required. If the file is opened with file_access.ReadWrite, both FileIOPermissionAccess.Read and FileIOPermissionAccess.Write permissions are required.
    /// </summary>
    open_or_create,
    /// <summary>
    /// Specifies that the operating system should open an existing file. When the file is opened, it should be truncated so that its size is zero bytes. This requires FileIOPermissionAccess.Write permission. Attempts to read from a file opened with file_mode.Truncate cause an ArgumentException exception.
    /// </summary>
    truncate
};

/// <summary>
/// Specifies whether the target of a relationship is inside or outside the Package.
/// </summary>
enum class target_mode
{
    /// <summary>
    /// The relationship references a part that is inside the package.
    /// </summary>
    external,
    /// <summary>
    /// The relationship references a resource that is external to the package.
    /// </summary>
    internal
};

enum class encoding_type
{
    utf8,
    latin1
};

class bad_sheet_title : public std::runtime_error
{
public:
    bad_sheet_title(const std::string &title) 
        : std::runtime_error(std::string("bad worksheet title: ") + title)
    {

    }
};

class bad_cell_coordinates : public std::runtime_error
{
public:
    bad_cell_coordinates(int row, int column)
        : std::runtime_error(std::string("bad cell coordinates: (") + std::to_string(row) + "," + std::to_string(column) + ")")
    {

    }
    
    bad_cell_coordinates(const std::string &coord_string)
    : std::runtime_error(std::string("bad cell coordinates: (") + coord_string + ")")
    {
        
    }
};

class zip_file
{
    enum class state
    {
        read,
        write,
        closed
    };

public:
    zip_file(const std::string &filename, file_mode mode, file_access access = file_access::read);

    ~zip_file();

    std::string get_file_contents(const std::string &filename);

    void set_file_contents(const std::string &filename, const std::string &contents);

    void delete_file(const std::string &filename);

    bool has_file(const std::string &filename);

    void flush(bool force_write = false);

private:
    void read_all();
    void write_all();
    std::string read_from_zip(const std::string &filename);
    void write_to_zip(const std::string &filename, const std::string &content, bool append = true);
    void change_state(state new_state, bool append = true);
    static bool file_exists(const std::string& name);
    void start_read();
    void stop_read();
    void start_write(bool append);
    void stop_write();

    zipFile zip_file_;
    unzFile unzip_file_;
    state current_state_;
    std::string filename_;
    std::unordered_map<std::string, std::string> files_;
    bool modified_;
//    file_mode mode_;
    file_access access_;
    std::vector<std::string> directories_;
};

/// <summary>
/// Represents a value type that may or may not have an assigned value.
/// </summary>
template<typename T>
struct nullable
{
public:
    /// <summary>
    /// Initializes a new instance of the nullable&lt;T&gt; structure with a "null" value.
    /// </summary>
    nullable() : has_value_(false) {}

    /// <summary>
    /// Initializes a new instance of the nullable&lt;T&gt; structure to the specified value.
    /// </summary>
    nullable(const T &value) : has_value_(true), value_(value) {}

    /// <summary>
    /// gets a value indicating whether the current nullable&lt;T&gt; object has a valid value of its underlying type.
    /// </summary>
    bool has_value() const { return has_value_; }

    /// <summary>
    /// gets the value of the current nullable&lt;T&gt; object if it has been assigned a valid underlying value.
    /// </summary>
    T get_value() const
    {
        if(!has_value())
        {
            throw std::runtime_error("invalid operation");
        }

        return value_;
    }

    /// <summary>
    /// Indicates whether the current nullable&lt;T&gt; object is equal to a specified object.
    /// </summary>
    bool equals(const nullable<T> &other) const
    {
        if(!has_value())
        {
            return !other.has_value();
        }

        if(other.has_value())
        {
            return get_value() == other.get_value();
        }

        return false;
    }

    /// <summary>
    /// Retrieves the value of the current nullable&lt;T&gt; object, or the object's default value.
    /// </summary>
    T get_value_or_default()
    {
        if(has_value())
        {
            return get_value();
        }

        return T();
    }

    /// <summary>
    /// Retrieves the value of the current nullable&lt;T&gt; object or the specified default value.
    /// </summary>
    T get_value_or_default(const T &default_)
    {
        if(has_value())
        {
            return get_value();
        }

        return default_;
    }

    /// <summary>
    /// Returns the value of a specified nullable&lt;T&gt; value.
    /// </summary>
    operator T() const
    {
        return get_value();
    }

    /// <summary>
    /// Indicates whether the current nullable&lt;T&gt; object is equal to a specified object.
    /// </summary>
    bool operator==(std::nullptr_t) const
    {
        return !has_value_;
    }

    /// <summary>
    /// Indicates whether the current nullable&lt;T&gt; object is not equal to a specified object.
    /// </summary>
    bool operator!=(std::nullptr_t) const
    {
        return has_value_;
    }

    /// <summary>
    /// Indicates whether the current nullable&lt;T&gt; object is equal to a specified object.
    /// </summary>
    bool operator==(T other) const
    {
        return has_value_ && other == value_;
    }

    /// <summary>
    /// Indicates whether the current nullable&lt;T&gt; object is not equal to a specified object.
    /// </summary>
    bool operator!=(T other) const
    {
        return has_value_ && other != value_;
    }

    /// <summary>
    /// Indicates whether the current nullable&lt;T&gt; object is equal to a specified object.
    /// </summary>
    bool operator==(const nullable<T> &other) const
    {
        return equals(other);
    }

    /// <summary>
    /// Indicates whether the current nullable&lt;T&gt; object is not equal to a specified object.
    /// </summary>
    bool operator!=(const nullable<T> &other) const
    {
        return !(*this == other);
    }

private:
    bool has_value_;
    T value_;
};

/// <summary>
/// Provides static methods for the creation, copying, deletion, moving, and opening of a single file, and aids in the creation of FileStream objects.
/// </summary>
class file
{
public:
    /// <summary>
    /// Copies an existing file to a new file. Overwriting a file of the same name is allowed.
    /// </summary>
    static void copy(const std::string &source, const std::string &destination, bool overwrite = false);

    /// <summary>
    /// Determines whether the specified file exists.
    /// </summary>
    static bool exists(const std::string &path);
};
    
class reader
{
public:
    static worksheet read_worksheet(std::istream &handle, workbook &wb, const std::string &title, const std::unordered_map<int, std::string> &);
};
    
class writer
{
public:
    static std::string write_worksheet(worksheet ws);
    static std::string write_worksheet(worksheet ws, const std::unordered_map<std::string, std::string> &string_table);
    static std::string write_theme();
};

/// <summary>
/// Represents an association between a source Package or part, and a target object which can be a part or external resource.
/// </summary>
class relationship
{
public:
    enum class type
    {
        hyperlink,
        drawing,
        worksheet,
        sharedStrings,
        styles,
        theme
    };

    relationship(const std::string &type, const std::string &r_id = "", const std::string &target_uri = "") : id_(r_id), source_uri_(""), target_uri_(target_uri)
    {
        if(type == "hyperlink")
        {
            type_ = type::hyperlink;
            target_mode_ = target_mode::external;
        }
    }

    /// <summary>
    /// gets a string that identifies the relationship.
    /// </summary>
    std::string get_id() const { return id_; }

    /// <summary>
    /// gets the URI of the package or part that owns the relationship.
    /// </summary>
    std::string get_source_uri() const { return source_uri_; }

    /// <summary>
    /// gets a value that indicates whether the target of the relationship is or External to the Package.
    /// </summary>
    target_mode get_target_mode() const { return target_mode_; }

    /// <summary>
    /// gets the URI of the target resource of the relationship.
    /// </summary>
    std::string get_target_uri() const { return target_uri_; }

private:
    relationship(const std::string &id, const std::string &relationship_type_, const std::string &source_uri, target_mode target_mode, const std::string &target_uri);
    //relationship &operator=(const relationship &rhs) = delete;

    type type_;
    std::string id_;
    std::string source_uri_;
    std::string target_uri_;
    target_mode target_mode_;
};

typedef std::vector<relationship> relationship_collection;

struct coordinate
{
    std::string column;
    int row;
};

/// <summary>
/// Alignment options for use in styles.
/// </summary>
struct alignment
{
    enum class horizontal_alignment
    {
        general,
        left,
        right,
        center,
        center_continuous,
        justify
    };

    enum class vertical_alignment
    {
        bottom,
        top,
        center,
        justify
    };

    horizontal_alignment horizontal = horizontal_alignment::general;
    vertical_alignment vertical = vertical_alignment::bottom;
    int text_rotation = 0;
    bool wrap_text = false;
    bool shrink_to_fit = false;
    int indent = 0;
};

class string_table
{
public:
    int operator[](const std::string &key) const;
private:
    friend class string_table_builder;
    std::vector<std::string> strings_;
};

class string_table_builder
{
public:
    void add(const std::string &string);
    string_table &get_table() { return table_; }
    const string_table &get_table() const { return table_; }
private:
    string_table table_;
};

class number_format
{
public:
    enum class format
    {
        general,
        text,
        number,
        number00,
        number_comma_separated1,
        number_comma_separated2,
        percentage,
        percentage00,
        date_yyyymmdd2,
        date_yyyymmdd,
        date_ddmmyyyy,
        date_dmyslash,
        date_dmyminus,
        date_dmminus,
        date_myminus,
        date_xlsx14,
        date_xlsx15,
        date_xlsx16,
        date_xlsx17,
        date_xlsx22,
        date_datetime,
        date_time1,
        date_time2,
        date_time3,
        date_time4,
        date_time5,
        date_time6,
        date_time7,
        date_time8,
        date_timedelta,
        date_yyyymmddslash,
        currency_usd_simple,
        currency_usd,
        currency_eur_simple
    };

    static const std::unordered_map<int, std::string> builtin_formats;

    static std::string builtin_format_code(int index);

    static bool is_date_format(const std::string &format);
    static bool is_builtin(const std::string &format);

    format get_format_code() const { return format_code_; }
    void set_format_code(format format_code) { format_code_ = format_code; }
    void set_format_code(const std::string &format_code) { custom_format_code_ = format_code; }

private:
    std::string custom_format_code_ = "";
    format format_code_ = format::general;
    int format_index_ = 0;
};

struct color
{
    static const color black;
    static const color white;
    static const color red;
    static const color darkred;
    static const color blue;
    static const color darkblue;
    static const color green;
    static const color darkgreen;
    static const color yellow;
    static const color darkyellow;

    color(int index)
    {
        this->index = index;
    }

    int index;
};

class font
{
    enum class underline
    {
        none,
        double_,
        double_accounting,
        single,
        single_accounting
    };

/*    std::string name = "Calibri";
    int size = 11;
    bool bold = false;
    bool italic = false;
    bool superscript = false;
    bool subscript = false;
    underline underline = underline::none;
    bool strikethrough = false;
    color color = color::black;*/
};

class fill
{
public:
    enum class type
    {
        none,
        solid,
        gradient_linear,
        gradient_path,
        pattern_darkdown,
        pattern_darkgray,
        pattern_darkgrid,
        pattern_darkhorizontal,
        pattern_darktrellis,
        pattern_darkup,
        pattern_darkvertical,
        pattern_gray0625,
        pattern_gray125,
        pattern_lightdown,
        pattern_lightgray,
        pattern_lightgrid,
        pattern_lighthorizontal,
        pattern_lighttrellis,
        pattern_lightup,
        pattern_lightvertical,
        pattern_mediumgray,
    };

    type type = type::none;
    int rotation = 0;
    color start_color = color::white;
    color end_color = color::black;
};

class borders
{
    struct border
    {
        enum class style
        {
            none,
            dashdot,
            dashdotdot,
            dashed,
            dotted,
            double_,
            hair,
            medium,
            mediumdashdot,
            mediumdashdotdot,
            mediumdashed,
            slantdashdot,
            thick,
            thin
        };

        style style = style::none;
        color color = color::black;
    };

    enum class diagonal_direction
    {
        none,
        up,
        down,
        both
    };

    border left;
    border right;
    border top;
    border bottom;
    border diagonal;
//    diagonal_direction diagonal_direction = diagonal_direction::none;
    border all_borders;
    border outline;
    border inside;
    border vertical;
    border horizontal;
};

class protection
{
public:
    enum class type
    {
        inherit,
        protected_,
        unprotected
    };

    type locked;
    type hidden;
};

class sheet_protection
{
public:
    static std::string hash_password(const std::string &password);

    void set_password(const std::string &password);
    std::string get_hashed_password() const { return hashed_password_; }

private:
    std::string hashed_password_;
};

class style
{
public:
    style(bool static_ = false) : static_(static_) {}

    style copy() const;

    font get_font() const;
    void set_font(font font);

    fill get_fill() const;
    void set_fill(fill fill);

    borders get_borders() const;
    void set_borders(borders borders);

    alignment get_alignment() const;
    void set_alignment(alignment alignment);

    number_format &get_number_format() { return number_format_; }
    const number_format &get_number_format() const { return number_format_; }
    void set_number_format(number_format number_format);

    protection get_protection() const;
    void set_protection(protection protection);

private:
    style(const style &rhs);

    bool static_ = false;
    font font_;
    fill fill_;
    borders borders_;
    alignment alignment_;
    number_format number_format_;
    protection protection_;
};

enum class optimized
{
    write,
    read,
    none
};

/// <summary>
/// Describes cell associated properties.
/// </summary>
/// <remarks>
/// Properties of interest include style, type, value, and address.
/// The Cell class is required to know its value and type, display options,
/// and any other features of an Excel cell.Utilities for referencing
/// cells using Excel's 'A1' column/row nomenclature are also provided.
/// </remarks>
class cell
{
public:
    enum class type
    {
        null,
        numeric,
        string,
        date,
        formula,
        boolean,
        error,
        hyperlink
    };

    static const std::unordered_map<std::string, int> ErrorCodes;

    /// <summary>
    /// Convert a coordinate string like 'B12' to a tuple ('B', 12)
    /// </summary>
    static coordinate coordinate_from_string(const std::string &address);

    /// <summary>
    /// Convert a coordinate to an absolute coordinate string (B12 -> $B$12)
    /// </summary>
    static std::string absolute_coordinate(const std::string &absolute_address);

    /// <summary>
    /// Convert a column letter into a column number (e.g. B -> 2)
    /// </summary>
    /// <remarks>
    /// Excel only supports 1 - 3 letter column names from A->ZZZ, so we
    /// restrict our column names to 1 - 3 characters, each in the range A - Z.
    /// </remarks>
    static int column_index_from_string(const std::string &column_string);

    /// <summary>
    /// Convert a column number into a column letter (3 -> 'C')
    /// </summary>
    /// <remarks>
    /// Right shift the column col_idx by 26 to find column letters in reverse
    /// order.These numbers are 1 - based, and can be converted to ASCII
    /// ordinals by adding 64.
    /// </remarks>
    static std::string get_column_letter(int column_index);

    static std::string check_string(const std::string &value);
    static std::string check_numeric(const std::string &value);
    static std::string check_error(const std::string &value);

    cell();
    cell(worksheet &ws, const std::string &column, int row);
    cell(worksheet &ws, const std::string &column, int row, const std::string &initial_value);

    std::string get_value() const;

    int get_row() const;
    std::string get_column() const;

    encoding_type get_encoding() const;
    std::string to_string() const;

    void set_explicit_value(const std::string &value, type data_type);
    type data_type_for_value(const std::string &value);

    bool bind_value();
    bool bind_value(int value);
    bool bind_value(double value);
    bool bind_value(const std::string &value);
    bool bind_value(const char *value);
    bool bind_value(bool value);
    bool bind_value(const tm &value);

    bool get_merged() const;
    void set_merged(bool);

    std::string get_hyperlink() const;
    void set_hyperlink(const std::string &value);

    std::string get_hyperlink_rel_id() const;

    void set_number_format(const std::string &format_code);

    bool has_style() const;

    style &get_style();
    const style &get_style() const;

    type get_data_type() const;

    std::string get_coordinate() const;
    std::string get_address() const;

    cell get_offset(int row, int column);

    bool is_date() const;

    std::pair<int, int> get_anchor() const;

    comment get_comment() const;
    void set_comment(comment comment);

    cell &operator=(const cell &rhs);
    cell &operator=(bool value);
    cell &operator=(int value);
    cell &operator=(double value);
    cell &operator=(const std::string &value);
    cell &operator=(const char *value);
    cell &operator=(const tm &value);

    bool operator==(std::nullptr_t) const;
    bool operator==(bool comparand) const;
    bool operator==(int comparand) const;
    bool operator==(double comparand) const;
    bool operator==(const std::string &comparand) const;
    bool operator==(const char *comparand) const;
    bool operator==(const tm &comparand) const;
    bool operator==(const cell &comparand) const { return root_ == comparand.root_; }

    friend bool operator==(std::nullptr_t, const cell &cell);
    friend bool operator==(bool comparand, const cell &cell);
    friend bool operator==(int comparand, const cell &cell);
    friend bool operator==(double comparand, const cell &cell);
    friend bool operator==(const std::string &comparand, const cell &cell);
    friend bool operator==(const char *comparand, const cell &cell);
    friend bool operator==(const tm &comparand, const cell &cell);

private:
    friend struct worksheet_struct;
    cell(cell_struct *root);
    cell_struct *root_;
};

inline std::ostream &operator<<(std::ostream &stream, const xlnt::cell &cell)
{
    return stream << cell.to_string();
}

typedef std::vector<std::vector<xlnt::cell>> range;

struct page_setup
{
    enum class Break
    {
        None = 0,
        Row = 1,
        Column = 2
    };

    enum class SheetState
    {
        Visible,
        Hidden,
        VeryHidden
    };

    enum class PaperSize
    {
        Letter = 1,
        LetterSmall = 2,
        Tabloid = 3,
        Ledger = 4,
        Legal = 5,
        Statement = 6,
        Executive = 7,
        A3 = 8,
        A4 = 9,
        A4Small = 10,
        A5 = 11
    };

    enum class Orientation
    {
        Portrait,
        Landscape
    };

    Break break_;
    SheetState sheet_state;
    PaperSize paper_size;
    Orientation orientation;
    bool fit_to_page;
    bool fit_to_height;
    bool fit_to_width;
};
    
struct margins
{
public:
    double get_top() const { return top_; }
    void set_top(double top) { top_ = top; }
    double get_left() const { return left_; }
    void set_left(double left) { left_ = left; }
    double get_bottom() const { return bottom_; }
    void set_bottom(double bottom) { bottom_ = bottom; }
    double get_right() const { return right_; }
    void set_right(double right) { right_ = right; }
    double get_header() const { return header_; }
    void set_header(double header) { header_ = header; }
    double get_footer() const { return footer_; }
    void set_footer(double footer) { footer_ = footer; }
    
private:
    double top_;
    double left_;
    double bottom_;
    double right_;
    double header_;
    double footer_;
};

class worksheet
{
public:
    worksheet();
    worksheet(workbook &parent);
    worksheet(worksheet_struct *root);

    void operator=(const worksheet &other);
    xlnt::cell operator[](const std::string &address);
    std::string to_string() const;
    workbook &get_parent() const;
    void garbage_collect();
    std::list<cell> get_cell_collection();
    std::string get_title() const;
    void set_title(const std::string &title);
    xlnt::cell get_freeze_panes() const;
    void set_freeze_panes(cell top_left_cell);
    void set_freeze_panes(const std::string &top_left_coordinate);
    void unfreeze_panes();
    xlnt::cell cell(const std::string &coordinate);
    xlnt::cell cell(int row, int column);
    int get_highest_row() const;
    int get_highest_column() const;
    std::string calculate_dimension() const;
    xlnt::range range(const std::string &range_string);
    xlnt::range range(const std::string &range_string, int row_offset, int column_offset);
    relationship create_relationship(const std::string &relationship_type, const std::string &target_uri);
    //void add_chart(chart chart);
    void merge_cells(const std::string &range_string);
    void merge_cells(int start_row, int start_column, int end_row, int end_column);
    void unmerge_cells(const std::string &range_string);
    void unmerge_cells(int start_row, int start_column, int end_row, int end_column);
    void append(const std::vector<std::string> &cells);
    void append(const std::unordered_map<std::string, std::string> &cells);
    void append(const std::unordered_map<int, std::string> &cells);
    xlnt::range rows() const;
    xlnt::range columns() const;
    bool operator==(const worksheet &other) const;
    bool operator!=(const worksheet &other) const;
    bool operator==(std::nullptr_t) const;
    bool operator!=(std::nullptr_t) const;
    std::vector<relationship> get_relationships();
    page_setup &get_page_setup();
    margins &get_page_margins();
    std::string get_auto_filter() const;
    void set_auto_filter(const std::string &range_string);
    void unset_auto_filter();
    bool has_auto_filter() const;

private:
    friend class workbook;
    friend class cell;
    worksheet_struct *root_;
};

class named_range
{
public:
    named_range() : parent_worksheet_(nullptr), range_string_("") {}
    named_range(worksheet ws, const std::string &range_string) : parent_worksheet_(ws), range_string_(range_string) {}
    std::string get_range_string() const { return range_string_; }
    worksheet get_scope() const { return parent_worksheet_; }
    void set_scope(worksheet scope) { parent_worksheet_ = scope; }
    bool operator==(const named_range &comparand) const;

private:
    worksheet parent_worksheet_;
    std::string range_string_;
};

class drawing
{
public:
    drawing();

private:
    friend class worksheet;
    drawing(drawing_struct *root);
    drawing_struct *root_;
};

class workbook
{
public:
    //constructors
    workbook(optimized optimized = optimized::none);

    //prevent copy and assignment
    workbook(const workbook &) = delete;
    const workbook &operator=(const workbook &) = delete;

    void read_workbook_settings(const std::string &xml_source);

    //getters
    worksheet get_active_sheet();
    bool get_optimized_write() const { return optimized_write_; }
    encoding_type get_encoding() const { return encoding_; }
    bool get_guess_types() const { return guess_types_; }
    bool get_data_only() const { return data_only_; }

    //create
    worksheet create_sheet();
    worksheet create_sheet(std::size_t index);
    worksheet create_sheet(const std::string &title);
    worksheet create_sheet(std::size_t index, const std::string &title);

    //add
    void add_sheet(worksheet worksheet);
    void add_sheet(worksheet worksheet, std::size_t index);

    //remove
    void remove_sheet(worksheet worksheet);

    //container operations
    worksheet get_sheet_by_name(const std::string &sheet_name);

    bool contains(const std::string &key) const;

    int get_index(worksheet worksheet);

    worksheet operator[](const std::string &name);
    worksheet operator[](int index);

    std::vector<worksheet>::iterator begin();
    std::vector<worksheet>::iterator end();

    std::vector<worksheet>::const_iterator begin() const;
    std::vector<worksheet>::const_iterator end() const;

    std::vector<worksheet>::const_iterator cbegin() const;
    std::vector<worksheet>::const_iterator cend() const;

    std::vector<std::string> get_sheet_names() const;

    //named ranges
    void create_named_range(const std::string &name, worksheet worksheet, const std::string &range_string);
    std::vector<named_range> get_named_ranges();
    void add_named_range(const std::string &name, named_range named_range);
    bool has_named_range(const std::string &name, worksheet ws) const;
    named_range get_named_range(const std::string &name, worksheet ws);
    void remove_named_range(named_range named_range);
    
    //serialization
    void save(const std::string &filename);
    void load(const std::string &filename);

    bool operator==(const workbook &rhs) const;

    std::unordered_map<std::string, std::pair<std::string, std::string>> get_root_relationships() const;
    std::unordered_map<std::string, std::pair<std::string, std::string>> get_relationships() const;

private:
    bool optimized_write_;
    bool optimized_read_;
    bool guess_types_;
    bool data_only_;
    int active_sheet_index_;
    encoding_type encoding_;
    std::vector<worksheet> worksheets_;
    std::unordered_map<std::string, named_range> named_ranges_;
    std::vector<relationship> relationships_;
    std::vector<drawing> drawings_;
};

} // namespace xlnt
