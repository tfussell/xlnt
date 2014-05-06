#pragma once

namespace xlnt {

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

} // namespace xlnt
