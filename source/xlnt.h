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

struct cell_struct;
struct worksheet_struct;
struct package_impl;

class style;
class worksheet;
class worksheet;
class cell;
class relationship;
class workbook;
class package;

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
    part(package_impl &package, const std::string &uri, opcContainer *container);

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
    /// <summary>
    /// Opens a package with a given IO stream, file mode, and file access setting.
    /// </summary>
    static package open(std::iostream &stream, file_mode package_mode, file_access package_access);

    /// <summary>
    /// Opens a package at a given path using a given file mode, file access, and file share setting.
    /// </summary>
    static package open(const std::string &path, file_mode package_mode = file_mode::OpenOrCreate,
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

    package(std::iostream &stream, file_mode package_mode, file_access package_access);
    package(const std::string &path, file_mode package_mode, file_access package_access);

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

    static coordinate coordinate_from_string(const std::string &address);
    static int column_index_from_string(const std::string &column_string);
    static std::string get_column_letter(int column_index);
    static std::string absolute_coordinate(const std::string &absolute_address);

    cell();
    cell(worksheet &ws, const std::string &column, int row);
    cell(worksheet &ws, const std::string &column, int row, const std::string &initial_value);

    cell &operator=(int value);
    cell &operator=(double value);
    cell &operator=(const std::string &value);
    cell &operator=(const char *value);
    cell &operator=(const tm &value);

    friend bool operator==(const std::string &comparand, const cell &cell);
    friend bool operator==(const char *comparand, const cell &cell);
    friend bool operator==(const tm &comparand, const cell &cell);

    std::string to_string() const;
    bool is_date() const;
    style &get_style();
    type get_data_type();

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

class worksheet
{
public:
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

    void operator=(const worksheet &other);
    cell operator[](const std::string &address);
    std::string to_string() const;
    workbook &get_parent() const;
    void garbage_collect();
    std::set<cell> get_cell_collection();
    std::string get_title() const;
    void set_title(const std::string &title);
    cell get_freeze_panes() const;
    void set_freeze_panes(cell top_left_cell);
    void set_freeze_panes(const std::string &top_left_coordinate);
    void unfreeze_panes();
    xlnt::cell cell(const std::string &coordinate);
    xlnt::cell cell(int row, int column);
    int get_highest_row() const;
    int get_highest_column() const;
    std::string calculate_dimension() const;
    range range(const std::string &range_string, int row_offset, int column_offset);
    relationship create_relationship(const std::string &relationship_type);
    //void add_chart(chart chart);
    void merge_cells(const std::string &range_string);
    void merge_cells(int start_row, int start_column, int end_row, int end_column);
    void unmerge_cells(const std::string &range_string);
    void unmerge_cells(int start_row, int start_column, int end_row, int end_column);
    void append(const std::vector<xlnt::cell> &cells);
    void append(const std::unordered_map<std::string, xlnt::cell> &cells);
    void append(const std::unordered_map<int, xlnt::cell> &cells);
    xlnt::range rows() const;
    xlnt::range columns() const;
    bool operator==(const worksheet &other) const;
    bool operator!=(const worksheet &other) const;
    bool operator==(std::nullptr_t) const;
    bool operator!=(std::nullptr_t) const;

private:
    friend class workbook;
    worksheet(worksheet_struct *root);
    worksheet_struct *root_;
};

class workbook
{
public:
    workbook();
    workbook(const workbook &) = delete;
    const workbook &operator=(const workbook &) = delete;

    worksheet get_sheet_by_name(const std::string &sheet_name);
    worksheet get_active();
    worksheet create_sheet();
    worksheet create_sheet(std::size_t index);
    std::vector<std::string> get_sheet_names() const;
    std::vector<worksheet>::iterator begin();
    std::vector<worksheet>::iterator end();
    worksheet operator[](const std::string &name);
    void save(const std::string &filename);

private:
    worksheet active_worksheet_;
    std::vector<worksheet> worksheets_;
};

} // namespace xlnt
