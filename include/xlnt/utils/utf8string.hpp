#include <cstdint>
#include <string>
#include <vector>

#include <utf8.h>

namespace xlnt {

class utf8string
{
public:
    static utf8string from_utf8(const std::string &s);
    static utf8string from_latin1(const std::string &s);
    static utf8string from_utf16(const std::string &s);
    static utf8string from_utf32(const std::string &s);

    static bool is_valid(const std::string &s);

    bool is_valid() const;

private:
    std::vector<std::uint8_t> bytes_;
};

} // namespace xlnt
