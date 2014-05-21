#pragma once

#include <string>
#include <vector>

namespace xlnt {
    
class string_table_builder;
    
class string_table
{
public:
    int operator[](const std::string &key) const;
private:
    friend class string_table_builder;
    std::vector<std::string> strings_;
};

class string_table_builder
{
public:
    void add(const std::string &string);
    string_table &get_table() { return table_; }
    const string_table &get_table() const { return table_; }
private:
    string_table table_;
};
    
} // namespace xlnt
