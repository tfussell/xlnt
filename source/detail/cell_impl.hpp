#pragma once

#include <cstdlib>

#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/comment.hpp>
#include <xlnt/cell/index_types.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/utils/time.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/styles/number_format.hpp>

#include "comment_impl.hpp"

namespace {

std::pair<bool, long double> cast_numeric(const std::string &s)
{
    const char *str = s.c_str();
    char *str_end = nullptr;
    auto result = std::strtold(str, &str_end);
    if (str_end != str + s.size()) return { false, 0 };
    return { true, result };
}

std::pair<bool, long double> cast_percentage(const std::string &s)
{
    if (s.back() == '%')
    {
        auto number = cast_numeric(s.substr(0, s.size() - 1));

        if (number.first)
        {
            return { true, number.second / 100 };
        }
    }

    return { false, 0 };
}

std::pair<bool, xlnt::time> cast_time(const std::string &s)
{
    xlnt::time result;

    try
    {
        auto last_colon = s.find_last_of(':');
        if (last_colon == std::string::npos) return { false, result };
        double seconds = std::stod(s.substr(last_colon + 1));
        result.second = static_cast<int>(seconds);
        result.microsecond = static_cast<int>((seconds - static_cast<double>(result.second)) * 1e6);

        auto first_colon = s.find_first_of(':');

        if (first_colon == last_colon)
        {
            auto decimal_pos = s.find('.');
            if (decimal_pos != std::string::npos)
            {
                result.minute = std::stoi(s.substr(0, first_colon));
            }
            else
            {
                result.hour = std::stoi(s.substr(0, first_colon));
                result.minute = result.second;
                result.second = 0;
            }
        }
        else
        {
            result.hour = std::stoi(s.substr(0, first_colon));
            result.minute = std::stoi(s.substr(first_colon + 1, last_colon - first_colon - 1));
        }
    }
    catch (std::invalid_argument)
    {
        return { false, result };
    }

    return { true, result };
}

} // namespace

namespace xlnt {

class style;

namespace detail {

struct worksheet_impl;

struct cell_impl
{
    cell_impl();
    cell_impl(column_t column, row_t row);
    cell_impl(worksheet_impl *parent, column_t column, row_t row);
    cell_impl(const cell_impl &rhs);
    cell_impl &operator=(const cell_impl &rhs);

    cell self();

    void set_string(const std::string &s, bool guess_types);

    cell::type type_;

    worksheet_impl *parent_;

    column_t column_;
    row_t row_;

    std::string value_string_;
    long double value_numeric_;

    std::string formula_;

    bool has_hyperlink_;
    relationship hyperlink_;

    bool is_merged_;

    bool has_style_;
    std::size_t style_id_;

    std::unique_ptr<comment_impl> comment_;
};

} // namespace detail
} // namespace xlnt
