#pragma once

#include <cstddef>

#include <xlnt/styles/conditional_format.hpp>
#include <xlnt/utils/optional.hpp>
#include <xlnt/worksheet/range_reference.hpp>

namespace xlnt {

class border;
class fill;
class font;

namespace detail {

struct stylesheet;
struct worksheet_impl;

struct conditional_format_impl
{
	stylesheet *parent;
	range_reference target_range;
	worksheet_impl *target_sheet;

	std::size_t priority;
	std::size_t differential_format_id;

	condition when;

	optional<std::size_t> border_id;
	optional<std::size_t> fill_id;
	optional<std::size_t> font_id;
};

} // namespace detail
} // namespace xlnt
