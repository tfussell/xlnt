#pragma once

#include <iterator>
#include <vector>

namespace xlnt {
namespace detail {

struct workbook_impl
{
    workbook_impl();
    workbook_impl &operator=(const workbook_impl &other)
    {
        active_sheet_index_ = other.active_sheet_index_;
        worksheets_.clear();
        std::copy(other.worksheets_.begin(), other.worksheets_.end(), back_inserter(worksheets_));
        relationships_.clear();
        std::copy(other.relationships_.begin(), other.relationships_.end(), std::back_inserter(relationships_));
        drawings_.clear();
        std::copy(other.drawings_.begin(), other.drawings_.end(), back_inserter(drawings_));
        properties_ = other.properties_;
        return *this;
    }
    workbook_impl(const workbook_impl &other) : active_sheet_index_(other.active_sheet_index_), worksheets_(other.worksheets_), relationships_(other.relationships_), drawings_(other.drawings_), properties_(other.properties_)
    {
        
    }
    //bool guess_types_;
    //bool data_only_;
    int active_sheet_index_;
    std::vector<worksheet_impl> worksheets_;
    std::vector<relationship> relationships_;
    std::vector<drawing> drawings_;
    document_properties properties_;
};

} // namespace detail
} // namespace xlnt
