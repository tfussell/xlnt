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

	optional<std::size_t> alignment;
	optional<std::size_t> border;
	optional<std::size_t> fill;
	optional<std::size_t> font;
	optional<std::size_t> number_format;
	optional<std::size_t> protection;

	optional<std::string> style;
};

} // namespace detail
} // namespace xlnt
