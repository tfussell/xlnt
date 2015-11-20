#include <xlnt/worksheet/page_margins.hpp>

namespace xlnt {

page_margins::page_margins() : default_(true), top_(0), left_(0), bottom_(0), right_(0), header_(0), footer_(0)
{
}

bool page_margins::is_default() const
{
    return default_;
}

double page_margins::get_top() const
{
    return top_;
}

void page_margins::set_top(double top)
{
    default_ = false;
    top_ = top;
}

double page_margins::get_left() const
{
    return left_;
}

void page_margins::set_left(double left)
{
    default_ = false;
    left_ = left;
}

double page_margins::get_bottom() const
{
    return bottom_;
}

void page_margins::set_bottom(double bottom)
{
    default_ = false;
    bottom_ = bottom;
}

double page_margins::get_right() const
{
    return right_;
}

void page_margins::set_right(double right)
{
    default_ = false;
    right_ = right;
}

double page_margins::get_header() const
{
    return header_;
}

void page_margins::set_header(double header)
{
    default_ = false;
    header_ = header;
}

double page_margins::get_footer() const
{
    return footer_;
}

void page_margins::set_footer(double footer)
{
    default_ = false;
    footer_ = footer;
}
    
} // namespace xlnt
