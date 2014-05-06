#pragma once

namespace xlnt {

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

} // namespace xlnt
