#include <xlnt/utils/utf8string.hpp>

namespace xlnt {

utf8string utf8string::from_utf8(const std::string &s)
{
    utf8string result;
    std::copy(s.begin(), s.end(), std::back_inserter(result.bytes_));

    return result;
}

utf8string utf8string::from_latin1(const std::string &s)
{
    utf8string result;
    std::copy(s.begin(), s.end(), std::back_inserter(result.bytes_));

    return result;
}

utf8string utf8string::from_utf16(const std::string &s)
{
    utf8string result;
    utf8::utf16to8(s.begin(), s.end(), std::back_inserter(result.bytes_));

    return result;
}

utf8string utf8string::from_utf32(const std::string &s)
{
    utf8string result;
    utf8::utf32to8(s.begin(), s.end(), std::back_inserter(result.bytes_));

    return result;
}

bool utf8string::is_valid(const std::string &s)
{
    return utf8::is_valid(s.begin(), s.end());
}

bool utf8string::is_valid() const
{
    return utf8::is_valid(bytes_.begin(), bytes_.end());
}

} // namespace xlnt
