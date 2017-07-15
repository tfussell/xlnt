#include <iostream>
#include <memory>

#include <xlnt/xlnt_config.hpp>

namespace arrow {
class Table;
}

namespace xlnt {

std::shared_ptr<arrow::Table> XLNT_API xlsx2arrow(std::istream &s);
void XLNT_API arrow2xlsx(std::shared_ptr<arrow::Table> &table, std::ostream &s);

} // namespace xlnt
