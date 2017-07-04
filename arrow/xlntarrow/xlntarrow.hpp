#include <iostream>
#include <arrow/api.h>
#include <xlnt/xlnt_config.hpp>

namespace xlnt {
namespace arrow {

void XLNT_API xlsx2arrow(std::istream &s, ::arrow::Table &table);
void XLNT_API arrow2xlsx(const ::arrow::Table &table, std::istream &s);

}
}
