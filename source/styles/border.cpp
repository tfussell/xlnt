#include <xlnt/styles/border.hpp>

namespace xlnt {

bool border::operator==(const border &other) const
{
    return hash() == other.hash();
}

std::size_t border::hash() const
{
    std::size_t seed = 0;

    hash_combine(seed, start_assigned);
    if (start_assigned) hash_combine(seed, start.hash());
    hash_combine(seed, end_assigned);
    if (end_assigned) hash_combine(seed, end.hash());
    hash_combine(seed, left_assigned);
    if (left_assigned) hash_combine(seed, left.hash());
    hash_combine(seed, right_assigned);
    if (right_assigned) hash_combine(seed, right.hash());
    hash_combine(seed, top_assigned);
    if (top_assigned) hash_combine(seed, top.hash());
    hash_combine(seed, bottom_assigned);
    if (bottom_assigned) hash_combine(seed, bottom.hash());
    hash_combine(seed, diagonal_assigned);
    if (diagonal_assigned) hash_combine(seed, diagonal.hash());
    hash_combine(seed, vertical_assigned);
    if (vertical_assigned) hash_combine(seed, vertical.hash());
    hash_combine(seed, horizontal_assigned);
    if (horizontal_assigned) hash_combine(seed, horizontal.hash());

    return seed;
}

} // namespace xlnt
