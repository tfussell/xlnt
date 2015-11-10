#pragma once

#include <string>

#include "xlnt_config.hpp"

namespace xlnt {

class XLNT_CLASS sheet_protection
{
  public:
    static std::string hash_password(const std::string &password);

    void set_password(const std::string &password);
    std::string get_hashed_password() const
    {
        return hashed_password_;
    }

  private:
    std::string hashed_password_;
};

} // namespace xlnt
