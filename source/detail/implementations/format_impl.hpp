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
    optional<bool> alignment_applied;

	optional<std::size_t> border_id;
    optional<bool> border_applied;

	optional<std::size_t> fill_id;
    optional<bool> fill_applied;

	optional<std::size_t> font_id;
    optional<bool> font_applied;

	optional<std::size_t> number_format_id;
    optional<bool> number_format_applied;

	optional<std::size_t> protection_id;
    optional<bool> protection_applied;

    bool pivot_button_ = false;
    bool quote_prefix_ = false;

	optional<std::string> style;

    std::size_t references = 0;

    XLNT_API friend bool operator==(const format_impl &left, const format_impl &right)
    {
        return left.parent == right.parent
            && left.alignment_id == right.alignment_id
            && left.alignment_applied == right.alignment_applied
            && left.border_id == right.border_id
            && left.border_applied == right.border_applied
            && left.fill_id == right.fill_id
            && left.fill_applied == right.fill_applied
            && left.font_id == right.font_id
            && left.font_applied == right.font_applied
            && left.number_format_id == right.number_format_id
            && left.number_format_applied == right.number_format_applied
            && left.protection_id == right.protection_id
            && left.protection_applied == right.protection_applied
            && left.pivot_button_ == right.pivot_button_
            && left.quote_prefix_ == right.quote_prefix_
            && left.style == right.style;
    }
};

} // namespace detail
} // namespace xlnt
