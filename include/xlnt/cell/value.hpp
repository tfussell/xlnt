#pragma once

#include <cstdint>
#include <string>

namespace xlnt {

struct date;
struct datetime;
struct time;
struct timedelta;

class value
{
public:
    enum class type
    {
        null,
        numeric,
        string,
        boolean,
        error
    };

    static value null();
    static value error(const std::string &error_string);

    value();
    value(value &&v);
    value(const value &v);
    explicit value(bool b);
	explicit value(std::int8_t i);
	explicit value(std::int16_t i);
	explicit value(std::int32_t i);
	explicit value(std::int64_t i);
	explicit value(std::uint8_t i);
	explicit value(std::uint16_t i);
	explicit value(std::uint32_t i);
	explicit value(std::uint64_t i);
#ifdef _WIN32
	explicit value(unsigned long i);
#endif
	explicit value(float f);
	explicit value(double d);
	explicit value(long double d);
	explicit value(const char *s);
	explicit value(const std::string &s);
	explicit value(const date &d);
	explicit value(const datetime &d);
	explicit value(const time &t);
	explicit value(const timedelta &t);

    template<typename T>
    T get() const;

    template<typename T>
    T as() const;

    type get_type() const;
    bool is(type t) const;

    bool is_integral() const;

    std::string to_string() const;    

    value &operator=(value other);

    bool operator==(const value &comparand) const;
    bool operator==(bool comparand) const;
	bool operator==(std::int8_t comparand) const;
	bool operator==(std::int16_t comparand) const;
	bool operator==(std::int32_t comparand) const;
	bool operator==(std::int64_t comparand) const;
	bool operator==(std::uint8_t comparand) const;
	bool operator==(std::uint16_t comparand) const;
	bool operator==(std::uint32_t comparand) const;
	bool operator==(std::uint64_t comparand) const;
#ifdef _WIN32
	bool operator==(unsigned long comparand) const;
#endif
	bool operator==(float comparand) const;
    bool operator==(double comparand) const;
	bool operator==(long double comparand) const;
    bool operator==(const std::string &comparand) const;
    bool operator==(const char *comparand) const;
    bool operator==(const date &comparand) const;
    bool operator==(const time &comparand) const;
    bool operator==(const datetime &comparand) const;
    bool operator==(const timedelta &comparand) const;

    friend bool operator==(bool comparand, const value &v);
	friend bool operator==(std::int8_t comparand, const value &v);
	friend bool operator==(std::int16_t comparand, const value &v);
	friend bool operator==(std::int32_t comparand, const value &v);
	friend bool operator==(std::int64_t comparand, const value &v);
	friend bool operator==(std::uint8_t comparand, const value &v);
	friend bool operator==(std::uint16_t comparand, const value &v);
	friend bool operator==(std::uint32_t comparand, const value &v);
	friend bool operator==(std::uint64_t comparand, const value &v);
#ifdef _WIN32
	friend bool operator==(unsigned long comparand, const value &v);
#endif
	friend bool operator==(float comparand, const value &v);
    friend bool operator==(double comparand, const value &v);
	friend bool operator==(long double comparand, const value &v);
    friend bool operator==(const std::string &comparand, const value &v);
    friend bool operator==(const char *comparand, const value &v);
    friend bool operator==(const date &comparand, const value &v);
    friend bool operator==(const time &comparand, const value &v);
    friend bool operator==(const datetime &comparand, const value &v);

    friend void swap(value &left, value &right);

private:
    type type_;
    std::string string_value_;
    long double numeric_value_;
};

} // namespace xlnt
