#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <iostream>
#include <iterator>
//TODO: make this conditional on XLNT_STD_STRING once std::sto* functions are replaced
#include <string>
#include <vector>

#include <utf8.h>

#include <xlnt/utils/string.hpp>

namespace {

template<typename T>
std::size_t string_length(const T *arr)
{
	std::size_t i = 0;

	while (arr[i] != T(0))
	{
		i++;
	}

	return i;
}

xlnt::string::code_point to_upper(xlnt::string::code_point p)
{
    if(p >= U'a' && p <= U'z')
    {
        return U'A' + p - U'a';
    }
    
    return p;
}

xlnt::string::code_point to_lower(xlnt::string::code_point p)
{
    if(p >= U'A' && p <= U'Z')
    {
        return U'a' + p - U'A';
    }
    
    return p;
}

}

namespace xlnt {

template<>
string string::from(std::int8_t i)
{
	return string(std::to_string(i).c_str());
}

template<>
string string::from(std::int16_t i)
{
	return string(std::to_string(i).c_str());
}

template<>
string string::from(std::int32_t i)
{
	return string(std::to_string(i).c_str());
}

template<>
string string::from(std::int64_t i)
{
	return string(std::to_string(i).c_str());
}

template<>
string string::from(std::uint8_t i)
{
	return string(std::to_string(i).c_str());
}

template<>
string string::from(std::uint16_t i)
{
	return string(std::to_string(i).c_str());
}

template<>
string string::from(std::uint32_t i)
{
	return string(std::to_string(i).c_str());
}

template<>
string string::from(std::uint64_t i)
{
	return string(std::to_string(i).c_str());
}

#ifndef _MSC_VER
template<>
string string::from(std::size_t i)
{
	return string(std::to_string(i).c_str());
}
#endif

template<>
string string::from(float i)
{
	return string(std::to_string(i).c_str());
}

template<>
string string::from(double i)
{
	return string(std::to_string(i).c_str());
}

template<>
string string::from(long double i)
{
	return string(std::to_string(i).c_str());
}

#ifdef XLNT_STD_STRING
template<>
string string::from(const std::string &s)
{
    return string(s.data());
}
#endif

xlnt::string::iterator::iterator(xlnt::string *parent, size_type index)
    : parent_(parent),
      index_(index)
{
}

xlnt::string::iterator::iterator(const xlnt::string::iterator &other)
    : parent_(other.parent_),
      index_(other.index_)
{
}

string::code_point xlnt::string::iterator::operator*()
{
    return parent_->at(index_);
}

bool xlnt::string::iterator::operator==(const iterator &other) const
{
    return parent_ == other.parent_ && index_ == other.index_;
}

string::difference_type string::iterator::operator-(const iterator &other) const
{
    return index_ - other.index_;
}

xlnt::string::const_iterator::const_iterator(const xlnt::string *parent, size_type index)
    : parent_(parent),
      index_(index)
{
}

xlnt::string::const_iterator::const_iterator(const xlnt::string::iterator &other)
    : parent_(other.parent_),
      index_(other.index_)
{
}

xlnt::string::const_iterator::const_iterator(const xlnt::string::const_iterator &other)
    : parent_(other.parent_),
      index_(other.index_)
{
}

string::difference_type string::const_iterator::operator-(const const_iterator &other) const
{
    return index_ - other.index_;
}

const string::code_point xlnt::string::const_iterator::operator*() const
{
    return parent_->at(index_);
}

bool xlnt::string::const_iterator::operator==(const const_iterator &other) const
{
    return parent_ == other.parent_ && index_ == other.index_;
}

xlnt::string::iterator &xlnt::string::iterator::operator--()
{
    return *this -= 1;
}

xlnt::string::const_iterator &xlnt::string::const_iterator::operator--()
{
    return *this -= 1;
}

xlnt::string::iterator &xlnt::string::iterator::operator++()
{
    return *this += 1;
}

xlnt::string::const_iterator &xlnt::string::const_iterator::operator++()
{
    return *this += 1;
}

string::const_iterator &string::const_iterator::operator+=(int offset)
{
    auto new_index = static_cast<int>(index_) + offset;
    new_index = std::max<int>(0, std::min<int>(static_cast<int>(parent_->length()), new_index));
    index_ = static_cast<std::size_t>(new_index);
    
    return *this;
}

xlnt::string::iterator xlnt::string::iterator::operator+(int offset)
{
    iterator copy = *this;
    
	auto new_index = static_cast<int>(index_) + offset;
    new_index = std::max<int>(0, std::min<int>(static_cast<int>(parent_->length()), new_index));
    copy.index_ = static_cast<std::size_t>(new_index);
    
    return copy;
}

xlnt::string::const_iterator xlnt::string::const_iterator::operator+(int offset)
{
    const_iterator copy = *this;
    
	auto new_index = static_cast<int>(index_) + offset;
    new_index = std::max<int>(0, std::min<int>(static_cast<int>(parent_->length()), new_index));
    copy.index_ = static_cast<std::size_t>(new_index);
    
    return copy;
}

string::string(size_type initial_size)
	: data_(new std::vector<char>),
      code_point_byte_offsets_(new std::unordered_map<size_type, size_type>)
{
	data_->resize(initial_size + 1);
	data_->back() = '\0';
    
    code_point_byte_offsets_->insert({0, 0});
}

string::string() : string(size_type(0))
{
}

string::string(string &&str) : string()
{
	swap(*this, str);
}

string::string(const string &str)
    : data_(new std::vector<char>(*str.data_)),
      code_point_byte_offsets_(new std::unordered_map<size_type, size_type>(*str.code_point_byte_offsets_))
{
}

string::string(const string &str, size_type offset) : string(str, offset, str.length() - offset)
{
}

string::string(const string &str, size_type offset, size_type len) : string()
{
    auto part = str.substr(offset, len);
    
    *data_ = *part.data_;
    *code_point_byte_offsets_ = *part.code_point_byte_offsets_;
}

#ifdef XLNT_STD_STRING
string::string(const std::string &str) : string(str.data())
{
}
#endif

string::string(const utf_mb_wide_string str) : string()
{
    auto iter = str;
    
    while(*iter != '\0')
    {
        append(*iter);
        iter++;
    }
}

string::string(const utf8_string str) : string()
{
    auto start = str;
    auto end = str;
    
    while(*end != '\0')
    {
        ++end;
    }
    
    auto iter = start;
    
    while(iter != end)
    {
        append((utf32_char)utf8::next(iter, end));
    }
}

string::string(const utf16_string str) : string()
{
    auto iter = str;
    
    while(*iter != '\0')
    {
        append(*iter);
        iter++;
    }
}

string::string(const utf32_string str) : string()
{
auto iter = str;
    
    while(*iter != '\0')
    {
        append(*iter);
        iter++;
    }
}

string::~string()
{
	delete data_;
    data_ = nullptr;
    delete code_point_byte_offsets_;
    code_point_byte_offsets_ = nullptr;
}

string string::to_upper() const
{
	string upper;
    
    for(auto c : *this)
    {
        upper.append(::to_upper(c));
    }
    
    return upper;
}

string string::to_lower() const
{

	string lower;
    
    for(auto c : *this)
    {
        lower.append(::to_lower(c));
    }
    
    return lower;
}

string string::substr(size_type offset) const
{
	return substr(offset, length() - offset);
}

string string::substr(size_type offset, size_type len) const
{
    if(len != npos && offset + len < length())
    {
      	return string(begin() + static_cast<int>(offset), begin() + static_cast<int>(offset + len));
    }
    
    return string(begin() + static_cast<int>(offset), end());
}

string::code_point string::back() const
{
	return at(length() - 1);
}

string::code_point string::front() const
{
	return at(0);
}

string::size_type string::find(code_point c) const
{
	return find(c, 0);
}

string::size_type string::find(char c) const
{
	return find(c, 0);
}

string::size_type string::find(const string &str) const
{
	return find(str, 0);
}

string::size_type string::find(code_point c, size_type offset) const
{
	auto iter = begin() + static_cast<int>(offset);

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

string::size_type string::find(char c, size_type offset) const
{
	return find(static_cast<code_point>(c), offset);
}

string::size_type string::find(const string &str, size_type offset) const
{
    while (offset < length() - str.length())
    {
        if (substr(offset, str.length()) == str)
        {
            return offset;
        }
        
        offset++;
    }
    
	return npos;
}

string::size_type string::find_last_of(code_point c) const
{
	return find_last_of(c, 0);
}

string::size_type string::find_last_of(char c) const
{
	return find_last_of(c, 0);
}

string::size_type string::find_last_of(const string &str) const
{
	return find_last_of(str, 0);
}

string::size_type string::find_last_of(code_point c, size_type offset) const
{
    auto stop = begin() + static_cast<int>(offset);
    auto iter = end() - 1;
    
    while (iter != stop)
    {
        if (*iter == c)
        {
            return iter - begin();
        }
        
        --iter;
    }
    
	return *stop == c ? offset : npos;
}
string::size_type string::find_last_of(char c, size_type offset) const
{
    return find_last_of(static_cast<code_point>(c), offset);
}

string::size_type string::find_last_of(const string &str, size_type offset) const
{
    auto stop = begin() + static_cast<int>(offset);
    auto iter = end() - 1;
    
    while (iter != stop)
    {
        if(str.find(*iter) != npos)
        {
            return iter - begin();
        }
        
        --iter;
    }
    
	return npos;
}

string::size_type string::find_first_of(const string &str) const
{
	return find_first_of(str, 0);
}

string::size_type string::find_first_of(const string &str, size_type offset) const
{
    auto iter = begin() + static_cast<int>(offset);
    
    while (iter != end())
    {
        if(str.find(*iter) != npos)
        {
            return iter - begin();
        }
        
        ++iter;
    }
    
	return npos;
}

string::size_type string::find_first_not_of(const string &str) const
{
	return find_first_not_of(str, 0);
}

string::size_type string::find_first_not_of(const string &str, size_type offset) const
{
    auto iter = begin() + static_cast<int>(offset);
    
    while (iter != end())
    {
        if(str.find(*iter) == npos)
        {
            return iter - begin();
        }
        
        ++iter;
    }
    
	return npos;
}

string::size_type string::find_last_not_of(const string &str) const
{
	return find_last_not_of(str, 0);
}

string::size_type string::find_last_not_of(const string &str, size_type offset) const
{
    auto stop = begin() + static_cast<int>(offset);
    auto iter = end() - 1;
    
    while (iter != stop)
    {
        if(str.find(*iter) == npos)
        {
            return iter - begin();
        }
        
        --iter;
    }
    
	return npos;
}

void string::clear()
{
	data_->clear();
	data_->push_back('\0');
    code_point_byte_offsets_->clear();
    code_point_byte_offsets_->insert({0, 0});
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

template<>
double string::to() const
{
    return std::stod(std::string(data()));
}

#ifdef XLNT_STD_STRING
template<>
std::string string::to() const
{
    return std::string(data());
}
#endif

int string::to_hex() const
{
	return 0;
}

void string::erase(size_type index)
{
    auto start = code_point_byte_offsets_->at(index);
    auto next_start = code_point_byte_offsets_->at(index + 1);
    
    data_->erase(data_->begin() + start, data_->begin() + next_start);
    
    auto code_point_bytes = next_start - start;
    
    for(size_type i = index + 1; i < length(); i++)
    {
        code_point_byte_offsets_->at(i) = code_point_byte_offsets_->at(i + 1) - code_point_bytes;
    }
    
    code_point_byte_offsets_->erase(code_point_byte_offsets_->find(length()));
}

string &string::operator=(string rhs)
{
	swap(*this, rhs);

	return *this;
}

string::iterator string::begin()
{
	return iterator(this, 0);
}

string::const_iterator string::cbegin() const
{
	return const_iterator(this, 0);
}

string::iterator string::end()
{
	return iterator(this, length());
}

string::const_iterator string::cend() const
{
	return const_iterator(this, length());
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

std::size_t string::num_bytes() const
{
	return data_->size();
}

void string::append(char c)
{
	data_->back() = c;
	data_->push_back('\0');
    
    code_point_byte_offsets_->insert({length() + 1, num_bytes() - 1});
}

void string::append(wchar_t c)
{
    if(c < 128)
    {
        append(static_cast<char>(c));
        return;
    }
    
	data_->pop_back();
    
    std::array<char, 4> utf8_encoded {{0}};
    auto end = utf8::utf16to8(&c, &c + 1, utf8_encoded.begin());
    std::copy(utf8_encoded.begin(), end, std::back_inserter(*data_));
    
	data_->push_back('\0');
    
    code_point_byte_offsets_->insert({length() + 1, num_bytes() - 1});
}

void string::append(char16_t c)
{
    if(c < 128)
    {
        append(static_cast<char>(c));
        return;
    }
    
	data_->pop_back();
    
    std::array<char, 4> utf8_encoded {{0}};
    auto end = utf8::utf16to8(&c, &c + 1, utf8_encoded.begin());
    std::copy(utf8_encoded.begin(), end, std::back_inserter(*data_));
    
	data_->push_back('\0');
    
    code_point_byte_offsets_->insert({length() + 1, num_bytes() - 1});
}

void string::append(code_point c)
{
    if(c < 128)
    {
        append(static_cast<char>(c));
        return;
    }
    
	data_->pop_back();
    
    std::array<char, 4> utf8_encoded {{0}};
    auto end = utf8::utf32to8(&c, &c + 1, utf8_encoded.begin());
    std::copy(utf8_encoded.begin(), end, std::back_inserter(*data_));

	data_->push_back('\0');
    
    code_point_byte_offsets_->insert({length() + 1, num_bytes() - 1});
}

void string::append(const string &str)
{
	for (auto c : str)
	{
		append(c);
	}
}

void string::replace(size_type index, utf32_char value)
{
    std::array<byte, 4> encoded = {{0}};
    auto encoded_end = utf8::utf32to8(&value, &value + 1, encoded.begin());
    auto encoded_len = encoded_end - encoded.begin();
    
    auto data_start = code_point_byte_offsets_->at(index);
    auto data_end = code_point_byte_offsets_->at(index + 1);
    
    auto previous_len = static_cast<int>(data_end - data_start);
    int difference = static_cast<int>(encoded_len) - previous_len;
    
    if(difference < 0)
    {
        data_->erase(data_->begin() + data_end + difference, data_->begin() + data_end);
    }
    else if(difference > 0)
    {
        data_->insert(data_->begin() + data_start, difference, '\0');
    }

    for(std::size_t i = index + 1; i < code_point_byte_offsets_->size(); i++)
    {
        code_point_byte_offsets_->at(i) += difference;
    }
    
    auto iter = encoded.begin();
    auto data_iter = data_->begin() + data_start;
    
    while(iter != encoded_end)
    {
        *data_iter = *iter;
        ++data_iter;
        ++iter;
    }
}

string::code_point string::at(size_type index)
{
    if(index == length())
    {
        return U'\0';
    }
    
    return utf8::peek_next(data_->begin() + code_point_byte_offsets_->at(index), data_->end());
}

const string::code_point string::at(size_type index) const
{
    if(index == length())
    {
        return U'\0';
    }
    
    return utf8::peek_next(data_->begin() + code_point_byte_offsets_->at(index), data_->end());
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
		if (*left_iter != *right_iter)
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

XLNT_FUNCTION void swap(string &left, string &right)
{
	using std::swap;

	swap(left.data_, right.data_);
    swap(left.code_point_byte_offsets_, right.code_point_byte_offsets_);
}

XLNT_FUNCTION std::ostream &operator<<(std::ostream &left, string &right)
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

XLNT_FUNCTION string operator+(const char *left, const string &right)
{
	return string(left) + right;
}

XLNT_FUNCTION bool string::operator<(const string &other) const
{
	return std::string(data()) < std::string(other.data());
}

string::size_type string::length() const
{
    return code_point_byte_offsets_->size() - 1;
}

} // namespace xlnt
