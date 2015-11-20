#include <xlnt/worksheet/cell_vector.hpp>
#include <xlnt/cell/cell.hpp>

namespace xlnt {

template <>
cell cell_vector::iterator::operator*()
{
    return ws_[current_cell_];
}

cell_vector::iterator cell_vector::begin()
{
    return iterator(ws_, ref_.get_top_left(), order_);
}

cell_vector::iterator cell_vector::end()
{
    if (order_ == major_order::row)
    {
        auto past_end = ref_.get_bottom_right();
        past_end.set_column_index(past_end.get_column_index() + 1);
        return iterator(ws_, past_end, order_);
    }

    auto past_end = ref_.get_bottom_right();
    past_end.set_row(past_end.get_row() + 1);
    return iterator(ws_, past_end, order_);
}

cell_vector::const_iterator cell_vector::cbegin() const
{
    return const_iterator(ws_, ref_.get_top_left(), order_);
}

cell_vector::const_iterator cell_vector::cend() const
{
    if (order_ == major_order::row)
    {
        auto past_end = ref_.get_bottom_right();
        past_end.set_column_index(past_end.get_column_index() + 1);
        return const_iterator(ws_, past_end, order_);
    }

    auto past_end = ref_.get_bottom_right();
    past_end.set_row(past_end.get_row() + 1);
    return const_iterator(ws_, past_end, order_);
}

cell cell_vector::operator[](std::size_t cell_index)
{
    return get_cell(cell_index);
}

std::size_t cell_vector::num_cells() const
{
    if (order_ == major_order::row)
    {
        return ref_.get_width() + 1;
    }

    return ref_.get_height() + 1;
}

cell_vector::cell_vector(worksheet ws, const range_reference &reference, major_order order)
    : ws_(ws), ref_(reference), order_(order)
{
}

cell cell_vector::front()
{
    if (order_ == major_order::row)
    {
        return get_cell(ref_.get_top_left().get_column().index);
    }

    return get_cell(ref_.get_top_left().get_row());
}

cell cell_vector::back()
{
    if (order_ == major_order::row)
    {
        return get_cell(ref_.get_bottom_right().get_column().index);
    }

    return get_cell(ref_.get_top_left().get_row());
}

cell cell_vector::get_cell(std::size_t index)
{
    if (order_ == major_order::row)
    {
        return ws_.get_cell(ref_.get_top_left().make_offset((int)index, 0));
    }

    return ws_.get_cell(ref_.get_top_left().make_offset(0, (int)index));
}

std::size_t cell_vector::length() const
{
    return num_cells();
}

cell_vector::const_iterator cell_vector::begin() const
{
    return cbegin();
}
cell_vector::const_iterator cell_vector::end() const
{
    return cend();
}

} // namespace xlnt
