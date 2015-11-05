#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <iostream>
#include <iterator>
#include <vector>

#include <utf8.h>

#include <xlnt/utils/string.hpp>

namespace {

template<typename T>
std::size_t count_bytes(const T *arr)
{
	std::size_t i = 0;

	while (arr[i] != T(0))
	{
		i++;
	}

	return i;
}

}

namespace xlnt {

bool string::code_point::operator==(char rhs) const
{
	return get() == rhs;
}

bool string::code_point::operator!=(char rhs) const
{
	return get() != rhs;
}

bool string::code_point::operator<(char rhs) const
{
	return char(get()) < rhs;
}

bool string::code_point::operator<=(char rhs) const
{
	return char(get()) <= rhs;
}

bool string::code_point::operator>(char rhs) const
{
	return char(get()) > rhs;
}

bool string::code_point::operator>=(char rhs) const
{
	return char(get()) >= rhs;
}

string string::from(std::int8_t i)
{
	return string(std::to_string(i).c_str());
}

string string::from(std::int16_t i)
{
	return string(std::to_string(i).c_str());
}

string string::from(std::int32_t i)
{
	return string(std::to_string(i).c_str());
}

string string::from(std::int64_t i)
{
	return string(std::to_string(i).c_str());
}

string string::from(std::uint8_t i)
{
	return string(std::to_string(i).c_str());
}

string string::from(std::uint16_t i)
{
	return string(std::to_string(i).c_str());
}

string string::from(std::uint32_t i)
{
	return string(std::to_string(i).c_str());
}

string string::from(std::uint64_t i)
{
	return string(std::to_string(i).c_str());
}

string string::from(float i)
{
	return string(std::to_string(i).c_str());
}

string string::from(double i)
{
	return string(std::to_string(i).c_str());
}

string string::from(long double i)
{
	return string(std::to_string(i).c_str());
}

string::code_point &string::code_point::operator=(string::utf32_char value)
{
	utf8::append(value, data_ + index_);

	return *this;
}

string::utf32_char string::code_point::get() const
{
	return utf8::peek_next(data_ + index_, data_ + end_);
}

template<>
xlnt::string::common_iterator<false> &xlnt::string::common_iterator<false>::operator--()
{
	auto iter = current_.data_ + current_.index_;
	utf8::prior(iter, current_.data_);
	current_.index_ = iter - current_.data_;

	return *this;
}

template<>
xlnt::string::common_iterator<true> &xlnt::string::common_iterator<true>::operator--()
{
	auto iter = current_.data_ + current_.index_;
	utf8::prior(iter, current_.data_);
	current_.index_ = iter - current_.data_;

	return *this;
}

template<>
xlnt::string::common_iterator<false> &xlnt::string::common_iterator<false>::operator++()
{
	auto iter = current_.data_ + current_.index_;
	utf8::advance(iter, 1, current_.data_ + current_.end_);
	current_.index_ = iter - current_.data_;

	return *this;
}

template<>
xlnt::string::common_iterator<true> &xlnt::string::common_iterator<true>::operator++()
{
	auto iter = current_.data_ + current_.index_;
	utf8::advance(iter, 1, current_.data_ + current_.end_);
	current_.index_ = iter - current_.data_;

	return *this;
}

string::string(size_type initial_size)
	: length_(0),
	  data_(new std::vector<char>)
{
	data_->resize(initial_size + 1);
	data_->back() = '\0';
}

string::string() : string(size_type(0))
{
}

string::string(string &&str) : string()
{
	swap(*this, str);
}

string::string(const string &str) : string(str, 0)
{
}

string::string(const string &str, size_type offset) : string(str, offset, str.length() - offset)
{
}

string::string(const string &str, size_type offset, size_type len) : string()
{
	*this = str.substr(offset, len);
}

string::string(const utf_mb_wide_string str) : string()
{
	data_->clear();
	utf8::utf16to8(str, str + count_bytes(str), std::back_inserter(*data_));
	data_->push_back('\0');

	length_ = utf8::distance(data(), data() + data_->size() - 1);
}

string::string(const utf8_string str) : string()
{
	data_->clear();
	std::copy(str, str + count_bytes(str), std::back_inserter(*data_));
	data_->push_back('\0');

	length_ = utf8::distance(data(), data() + data_->size() - 1);
}

string::string(const utf16_string str) : string()
{
	data_->clear();
	utf8::utf16to8(str, str + count_bytes(str), std::back_inserter(*data_));
	data_->push_back('\0');

	length_ = utf8::distance(data(), data() + data_->size() - 1);
}

string::string(const utf32_string str) : string()
{
	data_->clear();
	utf8::utf32to8(str, str + count_bytes(str), std::back_inserter(*data_));
	data_->push_back('\0');

	length_ = utf8::distance(data(), data() + data_->size() - 1);
}

string::~string()
{
	delete data_;
}

string string::to_upper() const
{
	return *this;
}

string string::to_lower() const
{
	return *this;
}

string string::substr(size_type offset) const
{
	return substr(offset, length() - offset);
}

string string::substr(size_type offset, size_type len) const
{
	auto iter = begin() + offset;
	size_type i = 0;
	string result;
	
	while (i++ < len && iter != end())
	{
		result.append(*iter);
		++iter;
	}

	return result;
}

string::code_point string::back()
{
	return *(begin() + length_ + -1);
}

const string::code_point string::back() const
{
	return *(begin() + length_ + -1);
}

string::code_point string::front()
{
	return *begin();
}

const string::code_point string::front() const
{
	return *begin();
}

string::size_type string::find(code_point c) const
{
	return find(c, 0);
}

string::size_type string::find(char c) const
{
	return find(c, 0);
}

string::size_type string::find(string str) const
{
	return find(str, 0);
}

string::size_type string::find(code_point c, size_type offset) const
{
	auto iter = begin() + offset;

	while (iter != end())
	{
		if ((*iter).get() == c.get())
		{
			return offset;
		}

		++iter;
		offset++;
	}

	return npos;
}

string::size_type string::find(char c, size_type offset) const
{
	auto iter = begin() + offset;

	while (iter != end())
	{
		if (*iter == c)
		{
			return offset;
		}

		++iter;
		offset++;
	}

	return npos;
}

string::size_type string::find(string str, size_type offset) const
{
	return 0;
}

string::size_type string::find_last_of(code_point c) const
{
	return find_last_of(c, 0);
}

string::size_type string::find_last_of(char c) const
{
	return find_last_of(c, 0);
}

string::size_type string::find_last_of(string str) const
{
	return find_last_of(str, 0);
}

string::size_type string::find_last_of(code_point c, size_type offset) const
{
	return 0;
}
string::size_type string::find_last_of(char c, size_type offset) const
{
	return 0;
}

string::size_type string::find_last_of(string str, size_type offset) const
{
	return 0;
}

string::size_type string::find_first_of(string str) const
{
	return find_first_of(str, 0);
}

string::size_type string::find_first_of(string str, size_type offset) const
{
	return 0;
}

string::size_type string::find_first_not_of(string str) const
{
	return find_first_not_of(str, 0);
}

string::size_type string::find_first_not_of(string str, size_type offset) const
{
	return 0;
}

string::size_type string::find_last_not_of(string str) const
{
	return find_last_not_of(str, 0);
}

string::size_type string::find_last_not_of(string str, size_type offset) const
{
	return 0;
}

void string::clear()
{
	length_ = 0;
	data_->clear();
	data_->push_back('\0');
}

template<>
std::size_t string::to() const
{
	return std::stoull(std::string(data()));
}

template<>
std::uint32_t string::to() const
{
	return std::stoul(std::string(data()));
}

template<>
std::int32_t string::to() const
{
	return std::stoi(std::string(data()));
}

template<>
long double string::to() const
{
	return std::stold(std::string(data()));
}

int string::to_hex() const
{
	return 0;
}

void string::remove(code_point iter)
{

}

string &string::operator=(string rhs)
{
	swap(*this, rhs);

	return *this;
}

string::iterator string::begin()
{
	return iterator(data(), 0, length_);
}

string::const_iterator string::cbegin() const
{
	return const_iterator(data(), 0, length_);
}

string::iterator string::end()
{
	return iterator(data(), length_, length_);
}

string::const_iterator string::cend() const
{
	return const_iterator(data(), length_, length_);
}

string::byte_pointer string::data()
{
	return &data_->front();
}

string::const_byte_pointer string::data() const
{
	return &data_->front();
}

std::size_t string::hash() const
{
	static std::hash<std::string> hasher;
	return hasher(std::string(data()));
}

std::size_t string::bytes() const
{
	return data_->size();
}

void string::append(char c)
{
	data_->back() = c;
	data_->push_back('\0');

	length_++;
}

void string::append(wchar_t c)
{
	data_->pop_back();
	utf8::utf16to8(&c, &c + 1, std::back_inserter(*data_));
	data_->push_back('\0');

	length_++;
}

void string::append(char16_t c)
{
	data_->pop_back();
	utf8::utf16to8(&c, &c + 1, std::back_inserter(*data_));
	data_->push_back('\0');

	length_++;
}

void string::append(char32_t c)
{
	data_->pop_back();
	utf8::utf32to8(&c, &c + 1, std::back_inserter(*data_));
	data_->push_back('\0');

	length_++;
}

void string::append(string str)
{
	for (auto c : str)
	{
		append(c);
	}
}

void string::append(code_point c)
{
	append(c.get());
}

string::code_point string::at(size_type index)
{
	return *(begin() + index);
}

const string::code_point string::at(size_type index) const
{
	return *(begin() + index);
}

bool string::operator==(const_byte_pointer str) const
{
	return *this == string(str);
}

bool string::operator==(const string &str) const
{
	if (length() != str.length()) return false;

	auto left_iter = begin();
	auto right_iter = str.begin();

	while (left_iter != end() && right_iter != str.end())
	{
		if ((*left_iter).get() != (*right_iter).get())
		{
			return false;
		}

		++left_iter;
		++right_iter;
	}

	return (left_iter == end()) != (right_iter == end());
}

string::code_point string::operator[](size_type index)
{
	return at(index);
}

const string::code_point string::operator[](size_type index) const
{
	return at(index);
}

string string::operator+(const string &rhs) const
{
	string copy(*this);
	copy.append(rhs);

	return copy;
}

string &string::operator+=(const string &rhs)
{
	append(rhs);

	return *this;
}

void swap(string &left, string &right)
{
	using std::swap;

	swap(left.length_, right.length_);
	swap(left.data_, right.data_);
}

std::ostream &operator<<(std::ostream &left, string &right)
{
	auto d = right.data();
	std::size_t i = 0;

	while (d[i] != '\0')
	{
		left << d[i];
		i++;
	}

	return left;
}

string operator+(const char *left, const string &right)
{
	return string(left) + right;
}

bool string::operator<(const string &other) const
{
	return std::string(data()) < std::string(other.data());
}

} // namespace xlnt
