#pragma once

#include <cstddef>

namespace xlnt {
namespace detail {

struct formatting_record
{
	std::size_t index;

	bool pivot_button;
	bool quote_prefix;

	bool apply_alignment;
	std::size_t alignment_id;

	bool apply_border;
	std::size_t border_id;

	bool apply_fill;
	std::size_t fill_id;

	bool apply_font;
	std::size_t font_id;

	bool apply_number_format;
	std::size_t number_format_id;

	bool apply_protection;
	std::size_t protection_id;
};

} // namespace detail
} // namespace xlnt
