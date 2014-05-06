#include "workbook.h"

namespace xlnt {

class workbook_impl
{
public:
    worksheet get_sheet_by_name(const std::string &name)
    {
        return worksheets_[name];
    }

    std::unordered_map<std::string, worksheet> worksheets_;
};

workbook::workbook() : impl_(nullptr)
{

}

worksheet workbook::get_sheet_by_name(const std::string &name)
{
    if(impl_ == nullptr)
    {
        return worksheet();
    }

    return impl_->get_sheet_by_name(name);
}

} // namespace xlnt
