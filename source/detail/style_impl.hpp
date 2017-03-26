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

    bool pivot_button_ = false;
    bool quote_prefix_ = false;
};

} // namespace detail
} // namespace xlnt
