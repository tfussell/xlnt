#include <string>
#include <unordered_map>
#include <vector>

#include "cell_impl.h"

namespace xlnt {

class workbook;

namespace detail {

struct worksheet_impl
{
    worksheet_impl(workbook *parent_workbook, const std::string &title)
    : parent_(parent_workbook), title_(title), freeze_panes_("A1")
    {
    }
    
    worksheet_impl(const worksheet_impl &other)
    {
        *this = other;
    }
    
    void operator=(const worksheet_impl &other)
    {
        parent_ = other.parent_;
        title_ = other.title_;
        freeze_panes_ = other.freeze_panes_;
        cell_map_ = other.cell_map_;
        relationships_ = other.relationships_;
        page_setup_ = other.page_setup_;
        auto_filter_ = other.auto_filter_;
        page_margins_ = other.page_margins_;
        merged_cells_ = other.merged_cells_;
        named_ranges_ = other.named_ranges_;
    }
    
    workbook *parent_;
    std::string title_;
    cell_reference freeze_panes_;
    std::unordered_map<int, std::unordered_map<int, cell_impl>> cell_map_;
    std::vector<relationship> relationships_;
    page_setup page_setup_;
    range_reference auto_filter_;
    margins page_margins_;
    std::vector<range_reference> merged_cells_;
    std::unordered_map<std::string, range_reference> named_ranges_;
};

} // namespace detail
} // namespace xlnt
