#include <algorithm>
#include <fstream>
#include <unordered_map>

#include "workbook.h"
#include "worksheet.h"
#include "packaging/package.h"

namespace xlnt {

struct workbook_struct
{
    workbook_struct(workbook *parent) : parent_(parent)
    {
        worksheets_.emplace_back(*parent, "Sheet1");
        active_name_ = "Sheet1";
    }

    worksheet get_sheet_by_name(const std::string &name)
    {
        auto match = std::find_if(worksheets_.begin(), worksheets_.end(), [=](const worksheet &w) { return w.get_title() == name; });
        return *match;
    }

    worksheet get_active()
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

    worksheet create_sheet()
    {
        worksheets_.emplace_back(*parent_, "Sheet");
        worksheets_.back().set_title("Sheet" + std::to_string(worksheets_.size()));
        return worksheets_.back();
    }

    worksheet create_sheet(std::size_t index)
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

workbook::workbook() : root_(new workbook_struct(this))
{
}

worksheet workbook::get_sheet_by_name(const std::string &name)
{
    if(root_ == nullptr)
    {
        throw std::runtime_error("workbook not initialized");
    }

    return root_->get_sheet_by_name(name);
}

worksheet workbook::get_active()
{
    return root_->get_active();
}

worksheet workbook::create_sheet()
{
    return root_->create_sheet();
}

worksheet workbook::create_sheet(std::size_t index)
{
    return root_->create_sheet(index);
}

std::vector<worksheet>::iterator workbook::begin()
{
    return root_->begin();
}

std::vector<worksheet>::iterator workbook::end()
{
    return root_->end();
}

std::vector<std::string> workbook::get_sheet_names() const
{
    return root_->get_sheet_names();
}

worksheet workbook::operator[](const std::string &name)
{
    return get_sheet_by_name(name);
}

void workbook::save(const std::string &filename)
{
    root_->save(filename);
}

} // namespace xlnt
