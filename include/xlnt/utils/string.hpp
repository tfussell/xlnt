#pragma once

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <iostream>
#include <iterator>
#include <vector>

#include "xlnt_config.hpp"

namespace xlnt {

class XLNT_CLASS string
{
public:
	using value_type = std::uint32_t;
	using reference = value_type &;
	using const_reference = value_type &;
	using pointer = value_type *;
	using const_pointer = const pointer;
	using byte = char;
	using byte_pointer = byte *;
	using const_byte_pointer = const byte *;
	using difference_type = std::ptrdiff_t;
	using size_type = std::size_t;

	using utf_mb_narrow_char = char;
	using utf_mb_narrow_string = utf_mb_narrow_char[];
	using utf_mb_wide_char = wchar_t;
	using utf_mb_wide_string = utf_mb_wide_char[];
	using utf8_char = char;
	using utf8_string = utf8_char[];
	using utf16_char = char16_t;
	using utf16_string = utf16_char[];
	using utf32_char = char32_t;
	using utf32_string = utf32_char[];

	class code_point
	{
	public:
		code_point(byte_pointer data, size_type index, size_type end)
			: data_(data),
			index_(index),
			end_(end)
		{
		}

		code_point(const_byte_pointer data, size_type index, size_type end)
			: data_(const_cast<byte_pointer>(data)),
			index_(index),
			end_(end)
		{
		}

		code_point &operator=(utf32_char value);

		code_point &operator=(const code_point &value)
		{
			return *this = value.get();
		}

		bool operator==(char rhs) const;
		bool operator!=(char rhs) const;
		bool operator<(char rhs) const;
		bool operator<=(char rhs) const;
		bool operator>(char rhs) const;
		bool operator>=(char rhs) const;

		utf32_char get() const;

		byte_pointer data_;
		size_type index_;
		size_type end_;
	};

	template <bool is_const = true>
	class common_iterator : public std::iterator<std::bidirectional_iterator_tag, value_type>
	{
	public:
		common_iterator(byte_pointer data, size_type index, size_type length)
			: current_(data, index, length)
		{
		}

		common_iterator(const_byte_pointer data, size_type index, size_type length)
			: current_(data, index, length)
		{
		}

		common_iterator(const common_iterator<false> &other)
			: current_(other.current_)
		{
		}

		common_iterator(const common_iterator<true> &other)
			: current_(other.current_)
		{
		}

		code_point &operator*()
		{
			return current_;
		}

		bool operator==(const common_iterator &other) const
		{
			return current_.data_ == other.current_.data_ && current_.index_ == other.current_.index_;
		}

		bool operator!=(const common_iterator &other) const
		{
			return !(*this == other);
		}

		difference_type operator-(const common_iterator &other) const
		{
			return current_.index_ - other.current_.index_;
		}

		common_iterator operator+(size_type offset)
		{
			common_iterator copy = *this;
			size_type end = std::max<size_type>(0, std::min<size_type>(current_.end_, current_.index_ + offset));

			while (copy.current_.index_ != end)
			{
				end < copy.current_.index_ ? --copy : ++copy;
			}

			return copy;
		}

		common_iterator &operator--();

		common_iterator operator--(int)
		{
			iterator old = *this;
			--*this;

			return old;
		}

		common_iterator &operator++();

		common_iterator operator++(int)
		{
			common_iterator old = *this;
			++*this;

			return old;
		}

		friend class common_iterator<true>;

	private:
		code_point current_;
	};

	using iterator = common_iterator<false>;
	using const_iterator = common_iterator<true>;
	using reverse_iterator = std::reverse_iterator<iterator>;
	using const_reverse_iterator = std::reverse_iterator<const_iterator>;

	static const size_type npos = -1;

	static string from(std::int8_t i);
	static string from(std::int16_t i);
	static string from(std::int32_t i);
	static string from(std::int64_t i);
	static string from(std::uint8_t i);
	static string from(std::uint16_t i);
	static string from(std::uint32_t i);
	static string from(std::uint64_t i);
	static string from(float i);
	static string from(double i);
	static string from(long double i);

	string();
	string(string &&str);
	string(const string &str);
	string(const string &str, size_type offset);
	string(const string &str, size_type offset, size_type len);

	string(const utf_mb_wide_string str);
	string(const utf8_string str);
	string(const utf16_string str);
	string(const utf32_string str);

	template <class InputIterator>
	string(InputIterator first, InputIterator last) : string()
	{
		while (first != last)
		{
			append(*first);
			++first;
		}
	}

	~string();

	size_type length() const { return length_; }
	size_type bytes() const;

	bool empty() const { return length() == 0; }

	string to_upper() const;
	string to_lower() const;

	string substr(size_type offset) const;
	string substr(size_type offset, size_type len) const;

	code_point back();
	const code_point back() const;
	code_point front();
	const code_point front() const;

	size_type find(code_point c) const;
	size_type find(char c) const;
	size_type find(string str) const;

	size_type find(code_point c, size_type offset) const;
	size_type find(char c, size_type offset) const;
	size_type find(string str, size_type offset) const;

	size_type find_last_of(code_point c) const;
	size_type find_last_of(char c) const;
	size_type find_last_of(string str) const;

	size_type find_last_of(code_point c, size_type offset) const;
	size_type find_last_of(char c, size_type offset) const;
	size_type find_last_of(string str, size_type offset) const;

	size_type find_first_of(string str) const;
	size_type find_first_of(string str, size_type offset) const;

	size_type find_first_not_of(string str) const;
	size_type find_first_not_of(string str, size_type offset) const;

	size_type find_last_not_of(string str) const;
	size_type find_last_not_of(string str, size_type offset) const;

	void clear();

	template<typename T>
	T to() const;

	int to_hex() const;

	void remove(code_point iter);

	iterator begin();
	const_iterator begin() const { return cbegin(); }
	const_iterator cbegin() const;
	iterator end();
	const_iterator end() const { return cend(); }
	const_iterator cend() const;

	byte_pointer data();
	const_byte_pointer data() const;

	std::size_t hash() const;

	string &operator=(string rhs);

	void append(char c);
	void append(wchar_t c);
	void append(char16_t c);
	void append(char32_t c);
	void append(string str);
	void append(code_point c);

	code_point at(size_type index);
	const code_point at(size_type index) const;

	bool operator==(const_byte_pointer str) const;
	bool operator!=(const_byte_pointer str) const { return !(*this == str); }

	bool operator==(const string &str) const;
	bool operator!=(const string &str) const { return !(*this == str); }

	code_point operator[](size_type index);
	const code_point operator[](size_type index) const;

	string operator+(const string &rhs) const;
	string &operator+=(const string &rhs);

	bool operator<(const string &other) const;

	friend void swap(string &left, string &right);
	friend std::ostream &operator<<(std::ostream &left, string &right);
	friend string operator+(const char *left, const string &right);

	friend bool operator==(const char *left, const string &right) { return right == left; }
	friend bool operator!=(const char *left, const string &right) { return right != left; }

private:
	explicit string(size_type initial_size);

	size_type length_;
	std::vector<byte> *data_;
};

} // namespace xlnt

namespace std {

template<>
struct hash<xlnt::string>
{
	std::size_t operator()(const xlnt::string &str) const
	{
		return str.hash();
	}
};

template<>
struct less<xlnt::string>
{
	std::size_t operator()(const xlnt::string &left, const xlnt::string &right) const
	{
		return left < right;
	}
};

} // namespace std
