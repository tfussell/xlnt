#pragma once

#include <xlnt/utils/string.hpp>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

class XLNT_CLASS sheet_protection
{
  public:
    static string hash_password(const string &password);

    void set_password(const string &password);
    string get_hashed_password() const
    {
        return hashed_password_;
    }

  private:
    string hashed_password_;
};

} // namespace xlnt
