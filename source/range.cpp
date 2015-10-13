#include <xlnt/worksheet/range.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace xlnt {

template<>
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
    if(order_ == major_order::row)
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
    if(order_ == major_order::row)
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
    if(order_ == major_order::row)
    {
        return ref_.get_width() + 1;
    }

    return ref_.get_height() + 1;
}

cell_vector::cell_vector(worksheet ws, const range_reference &reference, major_order order)
    : ws_(ws),
    ref_(reference),
    order_(order)
{
}

cell cell_vector::front()
{
    if(order_ == major_order::row)
    {
        return get_cell(ref_.get_top_left().get_column_index());
    }

    return get_cell(ref_.get_top_left().get_row());
}

cell cell_vector::back()
{
    if(order_ == major_order::row)
    {
        return get_cell(ref_.get_bottom_right().get_column_index());
    }

    return get_cell(ref_.get_top_left().get_row());
}

cell cell_vector::get_cell(std::size_t index)
{
    if(order_ == major_order::row)
    {
        return ws_.get_cell(ref_.get_top_left().make_offset((int)index, 0));
    }
 
    return ws_.get_cell(ref_.get_top_left().make_offset(0, (int)index));
  }

range::range(worksheet ws, const range_reference &reference, major_order order)
    : ws_(ws),
    ref_(reference),
    order_(order)
{

}

range::~range()
{
}

cell_vector range::operator[](std::size_t index)
{
    return get_vector(index);
}

range_reference range::get_reference() const
{
    return ref_;
}

std::size_t range::length() const
{
    if(order_ == major_order::row)
    {
        return ref_.get_bottom_right().get_row() - ref_.get_top_left().get_row() + 1;
    }
    
    return ref_.get_bottom_right().get_column_index() - ref_.get_top_left().get_column_index() + 1;
}

bool range::operator==(const range &comparand) const
{
    return ref_ == comparand.ref_
      && ws_ == comparand.ws_
      && order_ == comparand.order_;
}

cell_vector range::get_vector(std::size_t vector_index)
{
    if(order_ == major_order::row)
    {
        range_reference reference(ref_.get_top_left().get_column_index(), 
				  ref_.get_top_left().get_row() + (int)vector_index,
				  ref_.get_bottom_right().get_column_index(), 
				  ref_.get_top_left().get_row() + (int)vector_index);
	return cell_vector(ws_, reference, order_);
    }

    range_reference reference(ref_.get_top_left().get_column_index() + (int)vector_index, 
			      ref_.get_top_left().get_row(),
			      ref_.get_top_left().get_column_index() + (int)vector_index, 
			      ref_.get_bottom_right().get_row());
    return cell_vector(ws_, reference, order_);
}

bool range::contains(const cell_reference &ref)
{
    return ref_.get_top_left().get_column_index() <= ref.get_column_index() && ref_.get_bottom_right().get_column_index() >= ref.get_column_index() && ref_.get_top_left().get_row() <= ref.get_row() && ref_.get_bottom_right().get_row() >= ref.get_row();
}

cell range::get_cell(const cell_reference &ref)
{
    return (*this)[ref.get_row()][ref.get_column_index()];
}

range::iterator range::begin()
{
    if(order_ == major_order::row)
    {
        cell_reference top_right(ref_.get_bottom_right().get_column_index(), 
				 ref_.get_top_left().get_row());
	range_reference row_range(ref_.get_top_left(), top_right);
	return iterator(ws_, row_range, order_);
    }
    
    cell_reference bottom_left(ref_.get_top_left().get_column_index(), 
			     ref_.get_bottom_right().get_row());
    range_reference row_range(ref_.get_top_left(), bottom_left);
    return iterator(ws_, row_range, order_);
}

range::iterator range::end()
{
    if(order_ == major_order::row)
    {
        auto past_end_row_index = ref_.get_bottom_right().get_row() + 1;
	cell_reference bottom_left(ref_.get_top_left().get_column_index(), past_end_row_index);
	cell_reference bottom_right(ref_.get_bottom_right().get_column_index(), past_end_row_index);
	return iterator(ws_, range_reference(bottom_left, bottom_right), order_);
    }

    auto past_end_column_index = ref_.get_bottom_right().get_column_index() + 1;
    cell_reference top_right(past_end_column_index, ref_.get_top_left().get_row());
    cell_reference bottom_right(past_end_column_index, ref_.get_bottom_right().get_row());
    return iterator(ws_, range_reference(top_right, bottom_right), order_);
}

range::const_iterator range::cbegin() const
{
    if(order_ == major_order::row)
    {
        cell_reference top_right(ref_.get_bottom_right().get_column_index(), 
				 ref_.get_top_left().get_row());
	range_reference row_range(ref_.get_top_left(), top_right);
	return const_iterator(ws_, row_range, order_);
    }
    
    cell_reference bottom_left(ref_.get_top_left().get_column_index(), 
			     ref_.get_bottom_right().get_row());
    range_reference row_range(ref_.get_top_left(), bottom_left);
    return const_iterator(ws_, row_range, order_);
}

range::const_iterator range::cend() const
{
    if(order_ == major_order::row)
    {
        auto past_end_row_index = ref_.get_bottom_right().get_row() + 1;
	cell_reference bottom_left(ref_.get_top_left().get_column_index(), past_end_row_index);
	cell_reference bottom_right(ref_.get_bottom_right().get_column_index(), past_end_row_index);
	return const_iterator(ws_, range_reference(bottom_left, bottom_right), order_);
    }

    auto past_end_column_index = ref_.get_bottom_right().get_column_index() + 1;
    cell_reference top_right(past_end_column_index, ref_.get_top_left().get_row());
    cell_reference bottom_right(past_end_column_index, ref_.get_bottom_right().get_row());
    return const_iterator(ws_, range_reference(top_right, bottom_right), order_);
}

template<>
cell_vector range::iterator::operator*()
{
    if(order_ == major_order::row)
    {
        range_reference reference(range_.get_top_left().get_column_index(),
				  current_cell_.get_row(),
				  range_.get_bottom_right().get_column_index(),
				  current_cell_.get_row());
	return cell_vector(ws_, reference, order_);
    }

    range_reference reference(current_cell_.get_column_index(),
			      range_.get_top_left().get_row(),
			      current_cell_.get_column_index(),
			      range_.get_bottom_right().get_row());
    return cell_vector(ws_, reference, order_);
}

} // namespace xlnt
