#include <stdexcept>

#include <xlnt/cell/value.hpp>
#include <xlnt/common/datetime.hpp>

namespace xlnt {

value value::error(const std::string &error_string)
{
    value v(error_string);
    v.type_ = type::error;
    return v;
}

value::value() : type_(type::null), numeric_value_(0)
{
}

value::value(value &&v)
{
    swap(*this, v);
}

value::value(const value &v)
{
    type_ = v.type_;
    numeric_value_ = v.numeric_value_;
    string_value_ = v.string_value_;
}
    
value::value(bool b) : type_(type::boolean), numeric_value_(b ? 1 : 0)
{
}

value::value(std::int8_t i) : type_(type::numeric), numeric_value_(i)
{
}

value::value(std::int16_t i) : type_(type::numeric), numeric_value_(i)
{
}

value::value(std::int32_t i) : type_(type::numeric), numeric_value_(i)
{
}

value::value(std::int64_t i) : type_(type::numeric), numeric_value_(static_cast<long double>(i))
{
}

value::value(std::uint8_t i) : type_(type::numeric), numeric_value_(i)
{
}

value::value(std::uint16_t i) : type_(type::numeric), numeric_value_(i)
{
}

value::value(std::uint32_t i) : type_(type::numeric), numeric_value_(i)
{
}

value::value(std::uint64_t i) : type_(type::numeric), numeric_value_(static_cast<long double>(i))
{
}

#ifdef _WIN32
value::value(unsigned long i) : type_(type::numeric), numeric_value_(i)
{
}
#endif

#ifdef __linux__
value::value(long long i) : type_(type::numeric), numeric_value_(i)
{
}

value::value(unsigned long long i) : type_(type::numeric), numeric_value_(i)
{
}
#endif

value::value(float f) : type_(type::numeric), numeric_value_(f)
{
}

value::value(double d) : type_(type::numeric), numeric_value_(d)
{
}

value::value(long double d) : type_(type::numeric), numeric_value_(d)
{
}

value::value(const char *s) : value(std::string(s))
{
}

value::value(const std::string &s) : type_(type::string), string_value_(s)
{
}

value &value::operator=(value other)
{
    swap(*this, other);
    return *this;
}

bool value::is(type t) const
{
    return type_ == t;
}
    
template<>
std::string value::get() const
{
    if(type_ == type::string)
    {
        return string_value_;
    }
        
    throw std::runtime_error("not a string");
}

template<>
double value::as() const
{
    switch(type_)
    {
    case type::boolean:
    case type::numeric:
        return (double)numeric_value_;
    case type::string:
        return std::stod(string_value_);
    case type::error:
        throw std::runtime_error("invalid");
    case type::null:
        return 0;
    }

    return 0;
}

template<>
bool value::as() const
{
    switch(type_)
    {
    case type::boolean:
    case type::numeric:
        return numeric_value_ != 0;
    case type::string:
        return !string_value_.empty();
    case type::error:
        throw std::runtime_error("invalid");
    case type::null:
        return false;
    }

    return false;
}

template<>
std::int8_t value::as() const
{
	switch (type_)
	{
	case type::boolean:
	case type::numeric:
		return static_cast<std::int8_t>(numeric_value_);
	case type::string:
		return static_cast<std::int8_t>(std::stoi(string_value_));
	case type::error:
		throw std::runtime_error("invalid");
	case type::null:
		return 0;
	}

	return 0;
}

template<>
std::int16_t value::as() const
{
	switch (type_)
	{
	case type::boolean:
	case type::numeric:
		return static_cast<std::int16_t>(numeric_value_);
	case type::string:
		return static_cast<std::int16_t>(std::stoi(string_value_));
	case type::error:
		throw std::runtime_error("invalid");
	case type::null:
		return 0;
	}

	return 0;
}

template<>
std::int32_t value::as() const
{
	switch (type_)
	{
	case type::boolean:
	case type::numeric:
		return static_cast<std::int32_t>(numeric_value_);
	case type::string:
		return std::stoi(string_value_);
	case type::error:
		throw std::runtime_error("invalid");
	case type::null:
		return 0;
	}

	return 0;
}

template<>
std::int64_t value::as() const
{
	switch (type_)
	{
	case type::boolean:
	case type::numeric:
		return static_cast<std::int64_t>(numeric_value_);
	case type::string:
		return std::stoi(string_value_);
	case type::error:
		throw std::runtime_error("invalid");
	case type::null:
		return 0;
	}

	return 0;
}

template<>
std::uint8_t value::as() const
{
	switch (type_)
	{
	case type::boolean:
	case type::numeric:
		return static_cast<std::uint8_t>(numeric_value_);
	case type::string:
		return static_cast<std::uint8_t>(std::stoi(string_value_));
	case type::error:
		throw std::runtime_error("invalid");
	case type::null:
		return 0;
	}

	return 0;
}

template<>
std::uint16_t value::as() const
{
	switch (type_)
	{
	case type::boolean:
	case type::numeric:
		return static_cast<std::uint16_t>(numeric_value_);
	case type::string:
		return static_cast<std::uint16_t>(std::stoi(string_value_));
	case type::error:
		throw std::runtime_error("invalid");
	case type::null:
		return 0;
	}

	return 0;
}

template<>
std::uint32_t value::as() const
{
	switch (type_)
	{
	case type::boolean:
	case type::numeric:
		return static_cast<std::uint32_t>(numeric_value_);
	case type::string:
		return std::stoi(string_value_);
	case type::error:
		throw std::runtime_error("invalid");
	case type::null:
		return 0;
	}

	return 0;
}

template<>
std::uint64_t value::as() const
{
	switch (type_)
	{
	case type::boolean:
	case type::numeric:
		return static_cast<std::uint64_t>(numeric_value_);
	case type::string:
		return std::stoi(string_value_);
	case type::error:
		throw std::runtime_error("invalid");
	case type::null:
		return 0;
	}

	return 0;
}

#ifdef _WIN32
template<>
unsigned long value::as() const
{
	switch (type_)
	{
	case type::boolean:
	case type::numeric:
		return static_cast<unsigned long>(numeric_value_);
	case type::string:
		return std::stoi(string_value_);
	case type::error:
		throw std::runtime_error("invalid");
	case type::null:
		return 0;
	}

	return 0;
}
#endif

#ifdef __linux__
template<>
long long value::as() const
{
	switch (type_)
	{
	case type::boolean:
	case type::numeric:
		return static_cast<long long>(numeric_value_);
	case type::string:
		return std::stoi(string_value_);
	case type::error:
		throw std::runtime_error("invalid");
	case type::null:
		return 0;
	}

	return 0;
}

template<>
unsigned long long value::as() const
{
	switch (type_)
	{
	case type::boolean:
	case type::numeric:
		return static_cast<unsigned long long>(numeric_value_);
	case type::string:
		return std::stoi(string_value_);
	case type::error:
		throw std::runtime_error("invalid");
	case type::null:
		return 0;
	}

	return 0;
}
#endif

template<>
std::string value::as() const
{
	return to_string();
}

bool value::is_integral() const
{
    return type_ == type::numeric && (std::int64_t)numeric_value_ == numeric_value_;
}

value value::null()
{
    return value();
}

value::type value::get_type() const
{
    return type_;
}

std::string value::to_string() const
{
    switch(type_)
    {
    case type::boolean:
        return numeric_value_ != 0 ? "1" : "0";
    case type::numeric:
        return std::to_string(numeric_value_);
    case type::string:
    case type::error:
        return string_value_;
    case type::null:
        return "";
    }

    return "";
}

bool value::operator==(bool value) const
{
    return type_ == type::boolean && (numeric_value_ != 0) == value;
}

bool value::operator==(std::int8_t comparand) const
{
    return type_ == type::numeric && numeric_value_ == comparand;
}

bool value::operator==(std::int16_t comparand) const
{
	return type_ == type::numeric && numeric_value_ == comparand;
}

bool value::operator==(std::int32_t comparand) const
{
	return type_ == type::numeric && numeric_value_ == comparand;
}

bool value::operator==(std::int64_t comparand) const
{
	return type_ == type::numeric && numeric_value_ == comparand;
}

bool value::operator==(std::uint8_t comparand) const
{
	return type_ == type::numeric && numeric_value_ == comparand;
}

bool value::operator==(std::uint16_t comparand) const
{
	return type_ == type::numeric && numeric_value_ == comparand;
}

bool value::operator==(std::uint32_t comparand) const
{
	return type_ == type::numeric && numeric_value_ == comparand;
}

bool value::operator==(std::uint64_t comparand) const
{
	return type_ == type::numeric && numeric_value_ == comparand;
}

#ifdef _WIN32
bool value::operator==(unsigned long comparand) const
{
	return type_ == type::numeric && numeric_value_ == comparand;
}
#endif

bool value::operator==(float comparand) const
{
	return type_ == type::numeric && numeric_value_ == comparand;
}

bool value::operator==(double comparand) const
{
    return type_ == type::numeric && numeric_value_ == comparand;
}

bool value::operator==(long double comparand) const
{
	return type_ == type::numeric && numeric_value_ == comparand;
}

bool value::operator==(const std::string &comparand) const
{
    if(type_ == type::string)
    {
        return string_value_ == comparand;
    }

    return false;
}

bool value::operator==(const char *comparand) const
{
    return *this == std::string(comparand);
}

bool value::operator==(const time &comparand) const
{
    if(type_ != type::numeric)
    {
        return false;
    }

    return time::from_number(numeric_value_) == comparand;
}

bool value::operator==(const date &comparand) const
{
    return type_ == type::numeric && comparand.to_number(calendar::windows_1900) == numeric_value_;
}

bool value::operator==(const datetime &comparand) const
{
    return type_ == type::numeric && comparand.to_number(calendar::windows_1900) == numeric_value_;
}

bool value::operator==(const timedelta &comparand) const
{
    return type_ == type::numeric && comparand.to_number() == numeric_value_;
}

bool value::operator==(const value &v) const
{
    if(type_ != v.type_) return false;
    if(type_ == type::string || type_ == type::error) return string_value_ == v.string_value_;
    if(type_ == type::numeric || type_ == type::boolean) return numeric_value_ == v.numeric_value_;
    return true;
}

bool operator==(bool comparand, const xlnt::value &value)
{
    return value == comparand;
}

bool operator==(std::int8_t comparand, const xlnt::value &value)
{
    return value == comparand;
}

bool operator==(std::int16_t comparand, const xlnt::value &value)
{
	return value == comparand;
}

bool operator==(std::int32_t comparand, const xlnt::value &value)
{
	return value == comparand;
}

bool operator==(std::int64_t comparand, const xlnt::value &value)
{
	return value == comparand;
}

bool operator==(std::uint8_t comparand, const xlnt::value &value)
{
	return value == comparand;
}

bool operator==(std::uint16_t comparand, const xlnt::value &value)
{
	return value == comparand;
}

bool operator==(std::uint32_t comparand, const xlnt::value &value)
{
	return value == comparand;
}

bool operator==(std::uint64_t comparand, const xlnt::value &value)
{
	return value == comparand;
}

bool operator==(const char *comparand, const xlnt::value &value)
{
    return value == comparand;
}

#ifdef _WIN32
bool operator==(unsigned long comparand, const xlnt::value &value)
{
	return value == comparand;
}
#endif

bool operator==(const std::string &comparand, const xlnt::value &value)
{
    return value == comparand;
}

bool operator==(const time &comparand, const xlnt::value &value)
{
    return value == comparand;
}

bool operator==(const date &comparand, const xlnt::value &value)
{
    return value == comparand;
}

bool operator==(const datetime &comparand, const xlnt::value &value)
{
    return value == comparand;
}

void swap(value &left, value &right)
{
    using std::swap;
    swap(left.type_, right.type_);
    swap(left.string_value_, right.string_value_);
    swap(left.numeric_value_, right.numeric_value_);
}

} // namespace xlnt

