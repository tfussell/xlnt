#pragma once

#include <cstddef>
#include <cstdint>
#include <unordered_map>
#include <vector>

#ifdef XLNT_STD_STRING
#include <string>
#endif

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

template<typename T>
class iter_range
{
public:
    iter_range(T begin_iter, T end_iter) : begin_(begin_iter), end_(end_iter)
    {
    }
    
    T begin() { return begin_; }
    T end() { return end_; }
    
private:
    T begin_;
    T end_;
};

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

class XLNT_CLASS string
{
public:
    using code_point = utf32_char;
	using reference = code_point &;
	using const_reference = const code_point &;
	using pointer = code_point *;
	using const_pointer = const code_point *;
    using byte = char;
	using byte_pointer = byte *;
	using const_byte_pointer = const byte *;
	using difference_type = std::ptrdiff_t;
	using size_type = std::size_t;
    
    class const_iterator;
    
    class iterator : public std::iterator<std::bidirectional_iterator_tag, code_point>
	{
	public:
		iterator(string *parent, size_type index);
        
		iterator(const iterator &other);
        
		code_point operator*();
        
		bool operator==(const iterator &other) const;
        
		bool operator!=(const iterator &other) const { return !(*this == other); }

		difference_type operator-(const iterator &other) const;
        
        iterator &operator+=(int offset) { return *this = *this + offset; }
        
        iterator &operator-=(int offset) { return *this = *this - offset; }

		iterator operator+(int offset);
        
        iterator operator-(int offset) { return *this + (-1 * offset); }

		iterator &operator--();

		iterator operator--(int)
		{
			iterator old = *this;
			--*this;

			return old;
		}

		iterator &operator++();

		iterator operator++(int)
		{
			iterator old = *this;
			++*this;

			return old;
		}

	private:
        friend class const_iterator;
        
        string *parent_;
        size_type index_;
	};
    
	class const_iterator : public std::iterator<std::bidirectional_iterator_tag, code_point>
	{
	public:
		const_iterator(const string *parent, size_type index);
        
		const_iterator(const string::iterator &other);
        
		const_iterator(const const_iterator &other);
        
		const code_point operator*() const;
        
		bool operator==(const const_iterator &other) const;
        
		bool operator!=(const const_iterator &other) const { return !(*this == other); }

		difference_type operator-(const const_iterator &other) const;
        
        const_iterator &operator+=(int offset);
        
        const_iterator &operator-=(int offset) { return *this += -offset; }

		const_iterator operator+(int offset);
        
        const_iterator operator-(int offset) { return *this + (-1 * offset); }

		const_iterator &operator--();

		const_iterator operator--(int)
		{
			const_iterator old = *this;
			--*this;

			return old;
		}

		const_iterator &operator++();

		const_iterator operator++(int)
		{
			const_iterator old = *this;
			++*this;

			return old;
		}

	private:
        const string *parent_;
        size_type index_;
	};

	using reverse_iterator = std::reverse_iterator<iterator>;
	using const_reverse_iterator = std::reverse_iterator<const_iterator>;

	static const size_type npos = -1;

    template<typename T>
    static string from(T value);

	string();
	string(string &&str);
	string(const string &str);
	string(const string &str, size_type offset);
	string(const string &str, size_type offset, size_type len);
    
#ifdef XLNT_STD_STRING
    string(const std::string &str);
#endif

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

	size_type length() const;
	size_type num_bytes() const;

	bool empty() const { return length() == 0; }

	string to_upper() const;
	string to_lower() const;

	string substr(size_type offset) const;
	string substr(size_type offset, size_type len) const;

	code_point back() const;
	code_point front() const;

	size_type find(char c) const;
   	size_type find(code_point c) const;
	size_type find(const string &str) const;

	size_type find(char c, size_type offset) const;
	size_type find(code_point c, size_type offset) const;
	size_type find(const string &str, size_type offset) const;

	size_type find_last_of(char c) const;
	size_type find_last_of(code_point c) const;
	size_type find_last_of(const string &str) const;

	size_type find_last_of(char c, size_type offset) const;
	size_type find_last_of(code_point c, size_type offset) const;
	size_type find_last_of(const string &str, size_type offset) const;

	size_type find_first_of(const string &str) const;
	size_type find_first_of(const string &str, size_type offset) const;

	size_type find_first_not_of(const string &str) const;
	size_type find_first_not_of(const string &str, size_type offset) const;

	size_type find_last_not_of(const string &str) const;
	size_type find_last_not_of(const string &str, size_type offset) const;

	void clear();

	template<typename T>
	T to() const;

	int to_hex() const;

	void erase(size_type index);

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

	void append(utf_mb_narrow_char c);
	void append(utf_mb_wide_char c);
	void append(utf16_char c);
	void append(code_point c);
    void append(const string &str);
    
    void replace(size_type index, utf32_char c);

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

	std::vector<byte> *data_;
    std::unordered_map<size_type, size_type> *code_point_byte_offsets_;
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
