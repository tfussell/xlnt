#pragma once

namespace xlnt {

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

} // namespace xlnt
