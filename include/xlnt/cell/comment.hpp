#pragma once

#include <string>

namespace xlnt {

class comment
{
public:
    comment();
    comment(const std::string &text, const std::string &auth);
    ~comment();
    
    std::string get_text() const;
    std::string get_author() const;

private:
    std::string text_;
    std::string author_;
};

} // namespace xlnt
