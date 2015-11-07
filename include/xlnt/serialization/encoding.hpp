#pragma once

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

/// <summary>
/// Specifies the encoding used in a worksheet.
/// This isn't really supported yet.
/// </summary>
enum class XLNT_CLASS encoding
{
    utf8,
    latin1,
    utf16,
    utf32,
    ascii
};

} // namespace xlnt
