#include <stdexcept>

#include "common/string_table.hpp"

namespace xlnt {
    
int string_table::operator[](const std::string &key) const
{
    for(std::size_t i = 0; i < strings_.size(); i++)
    {
        if(key == strings_[i])
        {
            return (int)i;
        }
    }
    throw std::runtime_error("bad string");
}

void string_table_builder::add(const std::string &string)
{
    for(std::size_t i = 0; i < table_.strings_.size(); i++)
    {
        if(string == table_.strings_[i])
        {
            return;
        }
    }
    table_.strings_.push_back(string);
}
    
} // namespace xlnt
