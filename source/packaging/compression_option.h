#pragma once

namespace xlnt {

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

} // namespace xlnt
