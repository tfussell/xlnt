#pragma once

#include <cstdlib>

#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/comment.hpp>
#include <xlnt/cell/types.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/styles/number_format.hpp>
#include <xlnt/utils/datetime.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/utils/string.hpp>

#include <detail/comment_impl.hpp>

namespace {

// return s after checking encoding, size, and illegal characters
xlnt::string check_string(xlnt::string s)
{
    if (s.length() == 0)
    {
        return s;
    }

    // check encoding?

    if (s.length() > 32767)
    {
        s = s.substr(0, 32767); // max string length in Excel
    }

    for (xlnt::string::code_point c : s)
    {
        if (c >= 0 && (c <= 8 || c == 11 || c == 12 || (c >= 14 && c <= 31)))
        {
            throw xlnt::illegal_character_error(0);
        }
    }

    return s;
}

std::pair<bool, long double> cast_numeric(const xlnt::string &s)
{
    const char *str = s.data();
    char *str_end = nullptr;
    auto result = std::strtold(str, &str_end);
    if (str_end != str + s.length()) return { false, 0 };
    return { true, result };
}

std::pair<bool, long double> cast_percentage(const xlnt::string &s)
{
    if (s.back() == '%')
    {
        auto number = cast_numeric(s.substr(0, s.length() - 1));

        if (number.first)
        {
            return { true, number.second / 100 };
        }
    }

    return { false, 0 };
}

std::pair<bool, xlnt::time> cast_time(const xlnt::string &s)
{
    xlnt::time result;
    return { false, result };
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

    cell self()
    {
        return xlnt::cell(this);
    }

    void set_string(const string &s, bool guess_types)
    {
        value_string_ = check_string(s);
        type_ = cell::type::string;

        if (value_string_.length() > 1 && value_string_.front() == '=')
        {
            formula_ = value_string_;
            type_ = cell::type::formula;
            value_string_.length();
        }
        else if (cell::error_codes().find(s) != cell::error_codes().end())
        {
            type_ = cell::type::error;
        }
        else if (guess_types)
        {
            auto percentage = cast_percentage(s);

            if (percentage.first)
            {
                value_numeric_ = percentage.second;
                type_ = cell::type::numeric;
                self().set_number_format(xlnt::number_format::percentage());
            }
            else
            {
                auto time = cast_time(s);

                if (time.first)
                {
                    type_ = cell::type::numeric;
                    self().set_number_format(number_format::date_time6());
                    value_numeric_ = time.second.to_number();
                }
                else
                {
                    auto numeric = cast_numeric(s);

                    if (numeric.first)
                    {
                        value_numeric_ = numeric.second;
                        type_ = cell::type::numeric;
                    }
                }
            }
        }
    }

    cell::type type_;

    worksheet_impl *parent_;

    column_t column_;
    row_t row_;

    string value_string_;
    long double value_numeric_;

    string formula_;

    bool has_hyperlink_;
    relationship hyperlink_;

    bool is_merged_;

    bool has_style_;
    std::size_t style_id_;

    std::unique_ptr<comment_impl> comment_;
};

} // namespace detail
} // namespace xlnt
