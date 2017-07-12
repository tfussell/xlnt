#include <iostream>
#include <xlnt/xlnt_config.hpp>

namespace arrow {
class Table;
}

namespace xlnt {
namespace arrow {

void XLNT_API xlsx2arrow(std::istream &s, ::arrow::Table &table);
void XLNT_API arrow2xlsx(const ::arrow::Table &table, std::ostream &s);

} // namespace arrow
} // namespace xlnt
