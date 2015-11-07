#pragma once

#include <string>
#include <vector>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

class range_reference;
class worksheet;

std::vector<std::pair<string, string>> XLNT_FUNCTION split_named_range(const string &named_range_string);

class XLNT_CLASS named_range
{
  public:
    using target = std::pair<worksheet, range_reference>;

    named_range();
    named_range(const named_range &other);
    named_range(const string &name, const std::vector<target> &targets);

    string get_name() const
    {
        return name_;
    }
    const std::vector<target> &get_targets() const
    {
        return targets_;
    }

    named_range &operator=(const named_range &other);

  private:
    string name_;
    std::vector<target> targets_;
};
}
