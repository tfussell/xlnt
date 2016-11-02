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

    bool alignment_applied = false;
	optional<std::size_t> alignment_id;
    bool border_applied = false;
	optional<std::size_t> border_id;
    bool fill_applied = false;
	optional<std::size_t> fill_id;
    bool font_applied = false;
	optional<std::size_t> font_id;
    bool number_format_applied = false;
	optional<std::size_t> number_format_id;
    bool protection_applied = false;
	optional<std::size_t> protection_id;

	optional<std::string> style;
    
    bool operator==(const format_impl &other) const
    {
        return alignment_applied == other.alignment_applied
            && alignment_id.is_set() == other.alignment_id.is_set()
            && (!alignment_id.is_set() || alignment_id.get() == other.alignment_id.get())
            && border_applied == other.border_applied
            && border_id.is_set() == other.border_id.is_set()
            && (!border_id.is_set() || border_id.get() == other.border_id.get())
            && fill_applied == other.fill_applied
            && fill_id.is_set() == other.fill_id.is_set()
            && (!fill_id.is_set() || fill_id.get() == other.fill_id.get())
            && font_applied == other.font_applied
            && font_id.is_set() == other.font_id.is_set()
            && (!font_id.is_set() || font_id.get() == other.font_id.get())
            && number_format_applied == other.number_format_applied
            && number_format_id.is_set() == other.number_format_id.is_set()
            && (!number_format_id.is_set() || number_format_id.get() == other.number_format_id.get())
            && protection_applied == other.protection_applied
            && protection_id.is_set() == other.protection_id.is_set()
            && (!protection_id.is_set() || protection_id.get() == other.protection_id.get());
    }
};

} // namespace detail
} // namespace xlnt
