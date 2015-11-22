#include <xlnt/worksheet/worksheet.hpp>

#include "cell_impl.hpp"
#include "comment_impl.hpp"

namespace xlnt {
namespace detail {

cell_impl::cell_impl() : cell_impl(column_t(1), 1)
{
}

cell_impl::cell_impl(column_t column, row_t row) : cell_impl(nullptr, column, row)
{
}

cell_impl::cell_impl(worksheet_impl *parent, column_t column, row_t row)
    : type_(cell::type::null),
      parent_(parent),
      column_(column),
      row_(row),
      value_numeric_(0),
      has_hyperlink_(false),
      is_merged_(false),
      has_style_(false),
      style_id_(0),
      comment_(nullptr)
{
}

cell_impl::cell_impl(const cell_impl &rhs)
{
    *this = rhs;
}

cell_impl &cell_impl::operator=(const cell_impl &rhs)
{
    parent_ = rhs.parent_;
    value_numeric_ = rhs.value_numeric_;
    value_string_ = rhs.value_string_;
    hyperlink_ = rhs.hyperlink_;
    formula_ = rhs.formula_;
    column_ = rhs.column_;
    row_ = rhs.row_;
    is_merged_ = rhs.is_merged_;
    has_hyperlink_ = rhs.has_hyperlink_;
    type_ = rhs.type_;
    style_id_ = rhs.style_id_;
    has_style_ = rhs.has_style_;

    if (rhs.comment_ != nullptr)
    {
        comment_.reset(new comment_impl(*rhs.comment_));
    }

    return *this;
}

cell cell_impl::self()
{
    return xlnt::cell(this);
}

void cell_impl::set_string(const std::string &s, bool guess_types)
{
    value_string_ = s;
    type_ = cell::type::string;

    if (value_string_.size() > 1 && value_string_.front() == '=')
    {
        formula_ = value_string_;
        type_ = cell::type::formula;
        value_string_.clear();
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

} // namespace detail
} // namespace xlnt
