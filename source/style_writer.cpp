#include "writer/style_writer.hpp"

namespace xlnt {

style_writer::style_writer(xlnt::workbook &wb) : wb_(wb)
{
  
}

std::unordered_map<std::size_t, std::string> style_writer::get_style_by_hash() const
{
  return std::unordered_map<std::size_t, std::string>();
}

} // namespace xlnt
