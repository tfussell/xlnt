#pragma once

#include <ostream>
#include <set>
#include <string>
#include <unordered_map>
#include <vector>
#include <opc/opc.h>
#include <opc/container.h>

struct tm;

namespace xlnt {

class cell;
class comment;
class drawing;
class named_range;
class package;
class relationship;
class style;
class workbook;
class worksheet;

struct cell_struct;
struct drawing_struct;
struct named_range_struct;
struct package_impl;
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
    Read = 0x01,
    /// <summary>
    /// Read and write access to the file. Data can be written to and read from the file.
    /// </summary>
    ReadWrite = 0x02,
    /// <summary>
    /// Write access to the file. Data can be written to the file. Combine with Read for read/write access.
    /// </summary>
    Write = 0x04,
    /// <summary>
    /// Inherit access for part from enclosing package.
    /// </summary>
    FromPackage = 0x08
};

/// <summary>
/// Specifies how the operating system should open a file.
/// </summary>
enum class file_mode
{
    /// <summary>
    /// Opens the file if it exists and seeks to the end of the file, or creates a new file.This requires FileIOPermissionAccess.Append permission.file_mode.Append can be used only in conjunction with file_access.Write.Trying to seek to a position before the end of the file throws an IOException exception, and any attempt to read fails and throws a NotSupportedException exception.
    /// </summary>
    Append,
    /// <summary>
    /// Specifies that the operating system should create a new file. If the file already exists, it will be overwritten. This requires FileIOPermissionAccess.Write permission. file_mode.Create is equivalent to requesting that if the file does not exist, use CreateNew; otherwise, use Truncate. If the file already exists but is a hidden file, an UnauthorizedAccessException exception is thrown.
    /// </summary>
    Create,
    /// <summary>
    /// Specifies that the operating system should create a new file. This requires FileIOPermissionAccess.Write permission. If the file already exists, an IOException exception is thrown.
    /// </summary>
    CreateNew,
    /// <summary>
    /// Specifies that the operating system should open an existing file. The ability to open the file is dependent on the value specified by the file_access enumeration. A System.IO.FileNotFoundException exception is thrown if the file does not exist.
    /// </summary>
    Open,
    /// <summary>
    /// Specifies that the operating system should open a file if it exists; otherwise, a new file should be created. If the file is opened with file_access.Read, FileIOPermissionAccess.Read permission is required. If the file access is file_access.Write, FileIOPermissionAccess.Write permission is required. If the file is opened with file_access.ReadWrite, both FileIOPermissionAccess.Read and FileIOPermissionAccess.Write permissions are required.
    /// </summary>
    OpenOrCreate,
    /// <summary>
    /// Specifies that the operating system should open an existing file. When the file is opened, it should be truncated so that its size is zero bytes. This requires FileIOPermissionAccess.Write permission. Attempts to read from a file opened with file_mode.Truncate cause an ArgumentException exception.
    /// </summary>
    Truncate
};

/// <summary>
/// Contains constants for controlling the kind of access other FileStream objects can have to the same file.
/// </summary>
enum class file_share
{
    /// <summary>
    /// Allows subsequent deleting of a file.
    /// </summary>
    Delete,
    /// <summary>
    /// Makes the file handle inheritable by child processes. This is not directly supported by Win32.
    /// </summary>
    Inheritable,
    /// <summary>
    /// Declines sharing of the current file. Any request to open the file (by this process or another process) will 
    /// fail until the file is closed.
    /// </summary>
    None,
    /// <summary>
    /// Allows subsequent opening of the file for reading. If this flag is not specified, any request to open the file 
    /// for reading (by this process or another process) will fail until the file is closed. However, even if this flag
    /// is specified, additional permissions might still be needed to access the file.
    /// </summary>
    Read,
    /// <summary>
    /// Allows subsequent opening of the file for reading or writing. If this flag is not specified, any request to 
    /// open the file for reading or writing (by this process or another process) will fail until the file is closed. 
    /// However, even if this flag is specified, additional permissions might still be needed to access the file.
    /// </summary>
    ReadWrite,
    /// <summary>
    /// Allows subsequent opening of the file for writing. If this flag is not specified, any request to open the file
    /// for writing (by this process or another process) will fail until the file is closed. However, even if this flag
    /// is specified, additional permissions might still be needed to access the file.
    /// </summary>
    Write
};

/// <summary>
/// Specifies whether the target of a relationship is inside or outside the Package.
/// </summary>
enum class target_mode
{
    /// <summary>
    /// The relationship references a part that is inside the package.
    /// </summary>
    External,
    /// <summary>
    /// The relationship references a resource that is external to the package.
    /// </summary>
    Internal
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

    relationship(const std::string &type) : source_uri_(""), target_uri_("")
    {
        if(type == "hyperlink")
        {
            type_ = type::hyperlink;
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
    friend class package;

    relationship(const std::string &id, package &package, const std::string &relationship_type_, const std::string &source_uri, target_mode target_mode, const std::string &target_uri);
    relationship &operator=(const relationship &rhs) = delete;

    type type_;
    std::string id_;
    std::string source_uri_;
    std::string target_uri_;
    target_mode target_mode_;
};

typedef std::vector<relationship> relationship_collection;

class opc_callback_handler
{
public:
    static int read(void *context, char *buffer, int length);

    static int write(void *context, const char *buffer, int length);

    static int close(void *context);

    static opc_ofs_t seek(void *context, opc_ofs_t ofs);

    static int trim(void *context, opc_ofs_t new_size);

    static int flush(void *context);
};

struct part_struct;

/// <summary>
/// Represents a part that is stored in a ZipPackage.
/// </summary>
class part
{
public:
    /// <summary>
    /// Initializes a new instance of the part class with a specified parent Package, part URI, MIME content type, and compression_option.
    /// </summary>
    part(package_impl &package, const std::string &uri_part, const std::string &mime_type = "", compression_option compression = compression_option::NotCompressed);

    part &operator=(const part &) = delete;

    /// <summary>
    /// gets the compression option of the part content stream.
    /// </summary>
    compression_option get_compression_option() const;

    /// <summary>
    /// gets the MIME type of the content stream.
    /// </summary>
    std::string get_content_type() const;

    /// <summary>
    /// gets the parent Package of the part.
    /// </summary>
    package &get_package() const;

    /// <summary>
    /// gets the URI of the part.
    /// </summary>
    std::string get_uri() const;

    /// <summary>
    /// Creates a part-level relationship between this part to a specified target part or external resource.
    /// </summary>
    relationship create_relationship(const std::string &target_uri, target_mode target_mode, const std::string &relationship_type);

    /// <summary>
    /// Deletes a specified part-level relationship.
    /// </summary>
    void delete_relationship(const std::string &id);

    /// <summary>
    /// Returns the relationship that has a specified Id.
    /// </summary>
    relationship get_relationship(const std::string &id);

    /// <summary>
    /// Returns a collection of all the relationships that are owned by this part.
    /// </summary>
    relationship_collection get_relationships();

    /// <summary>
    /// Returns a collection of the relationships that match a specified RelationshipType.
    /// </summary>
    relationship_collection get_relationship_by_type(const std::string &relationship_type);

    /// <summary>
    /// Returns all the content of this part.
    /// </summary>
    std::string read();

    /// <summary>
    /// Writes the given data to the part stream.
    /// </summary>
    void write(const std::string &data);

    /// <summary>
    /// Returns a value that indicates whether this part owns a relationship with a specified Id.
    /// </summary>
    bool relationship_exists(const std::string &id) const;

    bool operator==(const part &comparand) const;
    bool operator==(const std::nullptr_t &) const;

private:
    friend struct package_impl;

    part(part_struct *root);

    part_struct *root_;
};

typedef std::vector<part> part_collection;

/// <summary>
/// Implements a derived subclass of the abstract Package base class—the ZipPackage class uses a 
/// ZIP archive as the container store. This class should not be inherited.
/// </summary>
class package
{
public:
    enum class type
    {
        Excel,
        Word,
        Powerpoint,
        Zip
    };

    package();
    ~package();

    type get_type() const;

    /// <summary>
    /// Opens a package with a given IO stream, file mode, and file access setting.
    /// </summary>
    void open(std::iostream &stream, file_mode package_mode, file_access package_access);

    /// <summary>
    /// Opens a package at a given path using a given file mode, file access, and file share setting.
    /// </summary>
    void open(const std::string &path, file_mode package_mode = file_mode::OpenOrCreate,
        file_access package_access = file_access::ReadWrite, file_share package_share = file_share::None);

    /// <summary>
    /// Saves and closes the package plus all underlying part streams.
    /// </summary>
    void close();

    /// <summary>
    /// Creates a new part with a given URI, content type, and compression option.
    /// </summary>
    part create_part(const std::string &part_uri, const std::string &content_type, compression_option compression = compression_option::Normal);

    /// <summary>
    /// Creates a package-level relationship to a part with a given URI, target mode, relationship type, and identifier (ID).
    /// </summary>
    relationship create_relationship(const std::string &part_uri, target_mode target_mode, const std::string &relationship_type, const std::string &id);

    /// <summary>
    /// Deletes a part with a given URI from the package.
    /// </summary>
    void delete_part(const std::string &part_uri);

    /// <summary>
    /// Deletes a package-level relationship.
    /// </summary>
    void delete_relationship(const std::string &id);

    /// <summary>
    /// Saves the contents of all parts and relationships that are contained in the package.
    /// </summary>
    void flush();

    /// <summary>
    /// Returns the part with a given URI.
    /// </summary>
    part get_part(const std::string &part_uri);

    /// <summary>
    /// Returns a collection of all the parts in the package.
    /// </summary>
    part_collection get_parts();

    /// <summary>
    /// 
    /// </summary>
    relationship get_relationship(const std::string &id);

    /// <summary>
    /// Returns the package-level relationship with a given identifier.
    /// </summary>
    relationship_collection get_relationships();

    /// <summary>
    /// Returns a collection of all the package-level relationships that match a given RelationshipType.
    /// </summary>
    relationship_collection get_relationships(const std::string &relationship_type);

    /// <summary>
    /// Indicates whether a part with a given URI is in the package.
    /// </summary>
    bool part_exists(const std::string &part_uri);

    /// <summary>
    /// Indicates whether a package-level relationship with a given ID is contained in the package.
    /// </summary>
    bool relationship_exists(const std::string &id);

    /// <summary>
    /// gets the file access setting for the package.
    /// </summary>
    file_access get_file_open_access() const;

    bool operator==(const package &comparand) const;
    bool operator==(const std::nullptr_t &) const;

    /// <summary>
    /// gets the category of the Package.
    /// </summary>
    std::string get_category() const;

    /// <summary>
    /// sets the category of the Package.
    /// </summary>
    void set_category(const std::string &category);

    /// <summary>
    /// gets the category of the Package.
    /// </summary>
    std::string get_content_status() const;

    /// <summary>
    /// sets the category of the Package.
    /// </summary>
    void set_content_status(const std::string &category);

    /// <summary>
    /// gets the category of the Package.
    /// </summary>
    std::string get_content_type() const;

    /// <summary>
    /// sets the category of the Package.
    /// </summary>
    void set_content_type(const std::string &category);

    /// <summary>
    /// gets the category of the Package.
    /// </summary>
    nullable<tm> get_created() const;

    /// <summary>
    /// sets the category of the Package.
    /// </summary>
    void set_created(const nullable<tm> &created);

    /// <summary>
    /// gets the category of the Package.
    /// </summary>
    std::string get_creator() const;

    /// <summary>
    /// sets the category of the Package.
    /// </summary>
    void set_creator(const std::string &category);

    /// <summary>
    /// gets the category of the Package.
    /// </summary>
    std::string get_description() const;

    /// <summary>
    /// sets the category of the Package.
    /// </summary>
    void set_description(const std::string &category);

    /// <summary>
    /// gets the category of the Package.
    /// </summary>
    std::string get_identifier() const;

    /// <summary>
    /// sets the category of the Package.
    /// </summary>
    void set_identifier(const std::string &category);

    /// <summary>
    /// gets the category of the Package.
    /// </summary>
    std::string get_keywords() const;

    /// <summary>
    /// sets the category of the Package.
    /// </summary>
    void set_keywords(const std::string &category);

    /// <summary>
    /// gets the category of the Package.
    /// </summary>
    std::string get_language() const;

    /// <summary>
    /// sets the category of the Package.
    /// </summary>
    void set_language(const std::string &category);

    /// <summary>
    /// gets the category of the Package.
    /// </summary>
    std::string get_last_modified_by() const;

    /// <summary>
    /// sets the category of the Package.
    /// </summary>
    void set_last_modified_by(const std::string &category);

    /// <summary>
    /// gets the category of the Package.
    /// </summary>
    nullable<tm> get_last_printed() const;

    /// <summary>
    /// sets the category of the Package.
    /// </summary>
    void set_last_printed(const nullable<tm> &created);

    /// <summary>
    /// gets the category of the Package.
    /// </summary>
    nullable<tm> get_modified() const;

    /// <summary>
    /// sets the category of the Package.
    /// </summary>
    void set_modified(const nullable<tm> &created);

    /// <summary>
    /// gets the category of the Package.
    /// </summary>
    std::string get_revision() const;

    /// <summary>
    /// sets the category of the Package.
    /// </summary>
    void set_revision(const std::string &category);

    /// <summary>
    /// gets the category of the Package.
    /// </summary>
    std::string getsubject() const;

    /// <summary>
    /// sets the category of the Package.
    /// </summary>
    void setsubject(const std::string &category);

    /// <summary>
    /// gets the category of the Package.
    /// </summary>
    std::string get_title() const;

    /// Ssummary>
    /// gets the category of the Package.
    /// </summary>
    void set_title(const std::string &category);

    /// <summary>
    /// gets the category of the Package.
    /// </summary>
    std::string get_version() const;

    /// <summary>
    /// sets the category of the Package.
    /// </summary>
    void set_version(const std::string &category);

private:
    friend opc_callback_handler;

    void open_container();

    int write(char *buffer, int length);

    int read(const char *buffer, int length);

    std::ios::pos_type seek(std::ios::pos_type ofs);

    int trim(std::ios::pos_type new_size);

    package_impl *impl_;
};

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

private:
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

    std::string name = "Calibri";
    int size = 11;
    bool bold = false;
    bool italic = false;
    bool superscript = false;
    bool subscript = false;
    underline underline = underline::none;
    bool strikethrough = false;
    color color = color::black;
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
    diagonal_direction diagonal_direction = diagonal_direction::none;
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
        error
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

inline std::ostream &operator<<(std::ostream &stream, const cell &cell)
{
    return stream << cell.to_string();
}

typedef std::vector<std::vector<cell>> range;

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

class worksheet
{
public:
    worksheet(workbook &parent);

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
    relationship create_relationship(const std::string &relationship_type);
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

private:
    friend class workbook;
    worksheet(worksheet_struct *root);
    worksheet_struct *root_;
};

class named_range
{
public:
    named_range(const std::string &name);

private:
    friend class worksheet;
    named_range(named_range_struct *root);
    named_range_struct *root_;
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

class writer
{
public:
    static std::string write_content_types(workbook &wb);
    static std::string write_root_rels(workbook &wb);
    static std::string write_worksheet(worksheet &ws);
    static std::string get_document_content(const std::string &filename);
};

class workbook
{
public:
    //constructors
    workbook();

    //prevent copy and assignment
    workbook(const workbook &) = delete;
    const workbook &operator=(const workbook &) = delete;

    //named parameters
    workbook &optimized_write(bool value);
    workbook &encoding(encoding_type value);
    workbook &guess_types(bool value);
    workbook &data_only(bool value);

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
    void add_named_range(named_range named_range);
    void get_named_range(const std::string &name);
    void remove_named_range(named_range named_range);
    
    //serialization
    void save(const std::string &filename);
    void load(const std::string &filename);

    bool operator==(const workbook &rhs) const;

private:
    bool optimized_write_;
    bool optimized_read_;
    bool guess_types_;
    bool data_only_;
    int active_sheet_index_;
    encoding_type encoding_;
    std::vector<worksheet> worksheets_;
    std::vector<named_range> named_ranges_;
    std::vector<relationship> relationships_;
    std::vector<drawing> drawings_;
};

} // namespace xlnt
