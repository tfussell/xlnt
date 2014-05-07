#include <algorithm>
#include <fstream>
#include <unordered_map>

#include "workbook.h"
#include "package.h"
#include "worksheet.h"

namespace xlnt {

class workbook_impl
{
public:
    workbook_impl(workbook *parent) : parent_(parent)
    {
        worksheets_.emplace_back(*parent, "Sheet1");
        active_name_ = "Sheet1";
    }

    worksheet &get_sheet_by_name(const std::string &name)
    {
        auto match = std::find_if(worksheets_.begin(), worksheets_.end(), [=](const worksheet &w) { return w.get_title() == name; });
        return *match;
    }

    worksheet &get_active()
    {
        return get_sheet_by_name(active_name_);
    }

    std::vector<worksheet>::iterator begin()
    {
        return worksheets_.begin();
    }

    std::vector<worksheet>::iterator end()
    {
        return worksheets_.end();
    }

    std::vector<std::string> get_sheet_names() const
    {
        std::vector<std::string> names;
        for(auto sheet : worksheets_)
        {
            names.push_back(sheet.get_title());
        }
        return names;
    }

    worksheet &create_sheet()
    {
        worksheets_.emplace_back(*parent_, "Sheet");
        worksheets_.back().set_title("Sheet" + std::to_string(worksheets_.size()));
        return worksheets_.back();
    }

    worksheet &create_sheet(std::size_t index)
    {
        worksheets_.emplace(worksheets_.begin() + index, *parent_, "Sheet");
        worksheets_[index].set_title("Sheet" + std::to_string(worksheets_.size()));
        return worksheets_[index];
    }

    void save(const std::string &filename)
    {
        auto package = package::open(filename);
        package.close();
    }

    std::vector<worksheet> worksheets_;
    std::string active_name_;
    workbook *parent_;
};

workbook::workbook() : impl_(nullptr)
{
    auto new_impl = std::make_shared<workbook_impl>(this);
    impl_.swap(new_impl);
}

worksheet &workbook::get_sheet_by_name(const std::string &name)
{
    if(impl_ == nullptr)
    {
        throw std::runtime_error("workbook not initialized");
    }

    return impl_->get_sheet_by_name(name);
}

worksheet &workbook::get_active()
{
    return impl_->get_active();
}

worksheet &workbook::create_sheet()
{
    return impl_->create_sheet();
}

worksheet &workbook::create_sheet(std::size_t index)
{
    return impl_->create_sheet(index);
}

std::vector<worksheet>::iterator workbook::begin()
{
    return impl_->begin();
}

std::vector<worksheet>::iterator workbook::end()
{
    return impl_->end();
}

std::vector<std::string> workbook::get_sheet_names() const
{
    return impl_->get_sheet_names();
}

worksheet &workbook::operator[](const std::string &name)
{
    return get_sheet_by_name(name);
}

void workbook::save(const std::string &filename)
{
    impl_->save(filename);
}

} // namespace xlnt
