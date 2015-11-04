
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/workbook/named_range.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace {

/// <summary>
/// Return a vector containing string split at each delim.
/// </summary>
/// <remark>
/// This should maybe be in a utility header so it can be used elsewhere.
/// </remarks>
std::vector<xlnt::string> split_string(const xlnt::string &string, char delim)
{
    std::vector<xlnt::string> split;
    xlnt::string::size_type previous_index = 0;
    auto separator_index = string.find(delim);

    while (separator_index != xlnt::string::npos)
    {
        auto part = string.substr(previous_index, separator_index - previous_index);
        split.push_back(part);

        previous_index = separator_index + 1;
        separator_index = string.find(delim, previous_index);
    }

    split.push_back(string.substr(previous_index));

    return split;
}

} // namespace

namespace xlnt {

std::vector<std::pair<string, string>> split_named_range(const string &named_range_string)
{
    std::vector<std::pair<string, string>> final;

    for (auto part : split_string(named_range_string, ','))
    {
        auto split = split_string(part, '!');

        if (split[0].front() == '\'' && split[0].back() == '\'')
        {
            split[0] = split[0].substr(1, split[0].length() - 2);
        }

        // Try to parse it. Use empty string if it's not a valid range.
        try
        {
            xlnt::range_reference ref(split[1]);
        }
        catch (xlnt::cell_coordinates_exception)
        {
            split[1] = "";
        }

        final.push_back({ split[0], split[1] });
    }

    return final;
}

named_range::named_range()
{
}

named_range::named_range(const named_range &other)
{
    *this = other;
}

named_range::named_range(const string &name, const std::vector<named_range::target> &targets)
    : name_(name), targets_(targets)
{
}

named_range &named_range::operator=(const named_range &other)
{
    name_ = other.name_;
    targets_ = other.targets_;

    return *this;
}

} // namespace xlnt
