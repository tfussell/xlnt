#include <xlnt/workbook/document_properties.hpp>

namespace xlnt {

document_properties::document_properties() 
    : created(1900, 1, 1), 
    modified(1900, 1, 1),
    excel_base_date(calendar::windows_1900)
{

}

} // namespace xlnt
