#pragma once

#include <string>
#include <vector>

namespace xlnt {

class range_reference;
class worksheet;

std::vector<std::pair<std::string, std::string>> split_named_range(const std::string &named_range_string);

class named_range
{
  public:
    using target = std::pair<worksheet, range_reference>;

    named_range();
    named_range(const named_range &other);
    named_range(const std::string &name, const std::vector<target> &targets);

    std::string get_name() const
    {
        return name_;
    }
    const std::vector<target> &get_targets() const
    {
        return targets_;
    }

    named_range &operator=(const named_range &other);

  private:
    std::string name_;
    std::vector<target> targets_;
};
}
