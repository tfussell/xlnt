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
        next_style_id_(other.next_style_id_),
        alignments_(other.alignments_),
    borders_(other.borders_),
    fills_(other.fills_),
    fonts_(other.fonts_),
    number_formats_(other.number_formats_),
    protections_(other.protections_),
    style_alignment_(other.style_alignment_),
    style_border_(other.style_border_),
    style_fill_(other.style_fill_),
    style_font_(other.style_font_),
    style_number_format_(other.style_number_format_),
    style_protection_(other.style_protection_),
    style_pivot_button_(other.style_pivot_button_),
    style_quote_prefix_(other.style_quote_prefix_)
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
        next_style_id_ = other.next_style_id_;
        alignments_ = other.alignments_;
        borders_ = other.borders_;
        fills_ = other.fills_;
        fonts_ = other.fonts_;
        number_formats_ = other.number_formats_;
        protections_ = other.protections_;
        style_alignment_ = other.style_alignment_;
        style_border_ = other.style_border_;
        style_fill_ = other.style_fill_;
        style_font_ = other.style_font_;
        style_number_format_ = other.style_number_format_;
        style_protection_ = other.style_protection_;
        style_pivot_button_ = other.style_pivot_button_;
        style_quote_prefix_ = other.style_quote_prefix_;
        
        return *this;
    }

    std::size_t active_sheet_index_;
    std::vector<worksheet_impl> worksheets_;
    std::vector<relationship> relationships_;
    std::vector<drawing> drawings_;
    document_properties properties_;
    bool guess_types_;
    bool data_only_;
    
    std::size_t next_style_id_;
    std::unordered_map<std::size_t, alignment> alignments_;
    std::unordered_map<std::size_t, border> borders_;
    std::unordered_map<std::size_t, fill> fills_;
    std::unordered_map<std::size_t, font> fonts_;
    std::unordered_map<std::size_t, number_format> number_formats_;
    std::unordered_map<std::size_t, protection> protections_;
    std::unordered_map<std::size_t, std::size_t> style_alignment_;
    std::unordered_map<std::size_t, std::size_t> style_border_;
    std::unordered_map<std::size_t, std::size_t> style_fill_;
    std::unordered_map<std::size_t, std::size_t> style_font_;
    std::unordered_map<std::size_t, std::size_t> style_number_format_;
    std::unordered_map<std::size_t, std::size_t> style_protection_;
    std::unordered_map<std::size_t, bool> style_pivot_button_;
    std::unordered_map<std::size_t, bool> style_quote_prefix_;
};

} // namespace detail
} // namespace xlnt
