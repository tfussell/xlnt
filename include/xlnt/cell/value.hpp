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
    value(bool b);
    value(int8_t i);
    value(int16_t i);
    value(int32_t i);
    value(int64_t i);
    value(uint8_t i);
    value(uint16_t i);
    value(uint32_t i);
    value(uint64_t i);
#ifdef _WIN32
    value(long long int i);
#endif
    value(float f);
    value(double d);
    value(long double d);
    value(const char *s);
    value(const std::string &s);
    value(const date &d);
    value(const datetime &d);
    value(const time &t);
    value(const timedelta &t);

    template<typename T>
    T get() const;

    template<typename T>
    T as() const;

    type get_type() const;
    bool is(type t) const;

    bool is_integral() const;

    std::string to_string() const;    

    value &operator=(value other);

    value &operator=(bool b);
    value &operator=(int8_t i);
    value &operator=(int16_t i);
    value &operator=(int32_t i);
    value &operator=(int64_t i);
    value &operator=(uint8_t i);
    value &operator=(uint16_t i);
    value &operator=(uint32_t i);
    value &operator=(uint64_t i);
    value &operator=(const char *s);
    value &operator=(const std::string &s);
    value &operator=(const date &d);
    value &operator=(const datetime &d);
    value &operator=(const time &t);
    value &operator=(const timedelta &t);

    bool operator==(const value &comparand) const;
    bool operator==(bool comparand) const;
    bool operator==(int comparand) const;
    bool operator==(double comparand) const;
    bool operator==(const std::string &comparand) const;
    bool operator==(const char *comparand) const;
    bool operator==(const date &comparand) const;
    bool operator==(const time &comparand) const;
    bool operator==(const datetime &comparand) const;
    bool operator==(const timedelta &comparand) const;

    friend bool operator==(bool comparand, const value &v);
    friend bool operator==(int comparand, const value &v);
    friend bool operator==(double comparand, const value &v);
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
