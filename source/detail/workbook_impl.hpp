#pragma once

#include <iterator>
#include <vector>

namespace xlnt {
namespace detail {

struct workbook_impl
{
    workbook_impl();
    
    workbook_impl(const workbook_impl &other) 
        : active_sheet_index_(other.active_sheet_index_),
        worksheets_(other.worksheets_), 
        relationships_(other.relationships_), 
        drawings_(other.drawings_), 
        properties_(other.properties_), 
        guess_types_(other.guess_types_),
        data_only_(other.data_only_),
        styles_(other.styles_),
        colors_(other.colors_),
        borders_(other.borders_),
        fills_(other.fills_),
        fonts_(other.fonts_),
        number_formats_(other.number_formats_)
    {
    }
    
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
        guess_types_ = other.guess_types_;
        data_only_ = other.data_only_;
        styles_ = other.styles_;
        borders_ = other.borders_;
        fills_ = other.fills_;
        fonts_ = other.fonts_;
        number_formats_ = other.number_formats_;
        colors_ = other.colors_;
        
        return *this;
    }

    std::size_t active_sheet_index_;
    std::vector<worksheet_impl> worksheets_;
    std::vector<relationship> relationships_;
    std::vector<drawing> drawings_;
    
    document_properties properties_;
    
    bool guess_types_;
    bool data_only_;
    
    std::vector<style> styles_;
    
    std::size_t next_custom_format_id_;
    
    std::vector<color> colors_;
    std::vector<border> borders_;
    std::vector<fill> fills_;
    std::vector<font> fonts_;
    std::vector<number_format> number_formats_;
    
};

} // namespace detail
} // namespace xlnt
