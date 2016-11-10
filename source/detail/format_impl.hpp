#pragma once

#include <cstddef>

#include <xlnt/styles/alignment.hpp>
#include <xlnt/styles/border.hpp>
#include <xlnt/styles/fill.hpp>
#include <xlnt/styles/font.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/styles/protection.hpp>
#include <xlnt/utils/optional.hpp>

namespace xlnt {

class alignment;
class border;
class fill;
class font;
class number_format;
class protection;

namespace detail {

struct stylesheet;

struct format_impl
{
	stylesheet *parent;

	std::size_t id;

	optional<std::size_t> alignment_id;
    bool alignment_applied = false;

	optional<std::size_t> border_id;
    bool border_applied = false;

	optional<std::size_t> fill_id;
    bool fill_applied = false;

	optional<std::size_t> font_id;
    bool font_applied = false;

	optional<std::size_t> number_format_id;
    bool number_format_applied = false;

	optional<std::size_t> protection_id;
    bool protection_applied = false;

	optional<std::string> style;
    
    std::size_t references = 0;
    
    XLNT_API friend bool operator==(const format_impl &left, const format_impl &right)
    {
        return left.parent == right.parent
            && left.alignment_id == right.alignment_id
            && left.border_id == right.border_id
            && left.fill_id == right.fill_id
            && left.font_id == right.font_id
            && left.number_format_id == right.number_format_id
            && left.protection_id == right.protection_id
            && left.style == right.style;
    }
};

} // namespace detail
} // namespace xlnt
