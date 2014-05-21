#pragma once

#include <string>
#include <unordered_map>
#include <vector>

#include <zip.h>
#include <unzip.h>

namespace xlnt {

/// <summary>
/// Defines constants for read, write, or read/write access to a file.
/// </summary>
enum class file_access
{
    /// <summary>
    /// Read access to the file. Data can only be read from the file. Combine with write for read/write access.
    /// </summary>
    read = 0x01,
    /// <summary>
    /// Read and write access to the file. Data can be both written to and read from the file.
    /// </summary>
    read_write = 0x03,
    /// <summary>
    /// Write access to the file. Data can only be written to the file. Combine with read for read/write access.
    /// </summary>
    write = 0x02
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
    
class zip_file
{
private:
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
    file_access access_;
    std::vector<std::string> directories_;
};
    
} // namespace xlnt
