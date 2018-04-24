#pragma once

#include <cstddef>
#include <string>

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

struct style_impl
{
	stylesheet *parent;

	std::string name;
	std::size_t formatting_record_id;

	bool custom_builtin;
	bool hidden_style;

	optional<std::size_t> builtin_id;
	optional<std::size_t> outline_style;

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
};

} // namespace detail
} // namespace xlnt
