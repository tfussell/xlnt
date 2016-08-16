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

	std::size_t formatting_record_id;

	optional<alignment> alignment;
	optional<border> border;
	optional<fill> fill;
	optional<font> font;
	optional<number_format> number_format;
	optional<protection> protection;
};

} // namespace detail
} // namespace xlnt
