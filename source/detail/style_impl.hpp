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

	optional<bool> custom_builtin;
	bool hidden_style;

	optional<std::size_t> builtin_id;
	optional<std::size_t> outline_style;

	optional<alignment> alignment;
	optional<border> border;
	optional<fill> fill;
	optional<font> font;
	optional<number_format> number_format;
	optional<protection> protection;
};

} // namespace detail
} // namespace xlnt
