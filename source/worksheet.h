#pragma once

#include <list>
#include <unordered_map>

#include "range.h"

namespace xlnt {
    
class cell;
class cell_reference;
class named_range;
class range_reference;
class relationship;
class workbook;
    
struct worksheet_struct;
    
struct page_setup
{
    enum class page_break
    {
        none = 0,
        row = 1,
        column = 2
    };
    
    enum class sheet_state
    {
        visible,
        hidden,
        very_hidden
    };
    
    enum class paper_size
    {
        letter = 1,
        letter_small = 2,
        tabloid = 3,
        ledger = 4,
        legal = 5,
        statement = 6,
        executive = 7,
        a3 = 8,
        a4 = 9,
        a4_small = 10,
        a5 = 11
    };
    
    enum class orientation
    {
        portrait,
        landscape
    };
    
public:
    page_setup() : default_(true), break_(page_break::none), sheet_state_(sheet_state::visible), paper_size_(paper_size::letter),
    orientation_(orientation::portrait), fit_to_page_(false), fit_to_height_(false), fit_to_width_(false) {}
    bool is_default() const { return default_; }
    page_break get_break() const { return break_; }
    void set_break(page_break b) { default_ = false; break_ = b; }
    sheet_state get_sheet_state() const { return sheet_state_; }
    void set_sheet_state(sheet_state sheet_state) { default_ = false; sheet_state_ = sheet_state; }
    paper_size get_paper_size() const { return paper_size_; }
    void set_paper_size(paper_size paper_size) { default_ = false; paper_size_ = paper_size; }
    orientation get_orientation() const { return orientation_; }
    void set_orientation(orientation orientation) { default_ = false; orientation_ = orientation; }
    bool fit_to_page() const { return fit_to_page_; }
    void set_fit_to_page(bool fit_to_page) { default_ = false; fit_to_page_ = fit_to_page; }
    bool fit_to_height() const { return fit_to_height_; }
    void set_fit_to_height(bool fit_to_height) { default_ = false; fit_to_height_ = fit_to_height; }
    bool fit_to_width() const { return fit_to_width_; }
    void set_fit_to_width(bool fit_to_width) { default_ = false; fit_to_width_ = fit_to_width; }
    
private:
    bool default_;
    page_break break_;
    sheet_state sheet_state_;
    paper_size paper_size_;
    orientation orientation_;
    bool fit_to_page_;
    bool fit_to_height_;
    bool fit_to_width_;
};

struct margins
{
public:
    margins() : default_(true), top_(0), left_(0), bottom_(0), right_(0), header_(0), footer_(0) {}
    
    bool is_default() const { return default_; }
    double get_top() const { return top_; }
    void set_top(double top) { default_ = false; top_ = top; }
    double get_left() const { return left_; }
    void set_left(double left) { default_ = false; left_ = left; }
    double get_bottom() const { return bottom_; }
    void set_bottom(double bottom) { default_ = false; bottom_ = bottom; }
    double get_right() const { return right_; }
    void set_right(double right) { default_ = false; right_ = right; }
    double get_header() const { return header_; }
    void set_header(double header) { default_ = false; header_ = header; }
    double get_footer() const { return footer_; }
    void set_footer(double footer) { default_ = false; footer_ = footer; }
    
private:
    bool default_;
    double top_;
    double left_;
    double bottom_;
    double right_;
    double header_;
    double footer_;
};

class worksheet
{
public:
    worksheet();
    worksheet(workbook &parent);
    worksheet(worksheet_struct *root);

    std::string to_string() const;
    workbook &get_parent() const;
    void garbage_collect();
    
    // title
    std::string get_title() const;
    void set_title(const std::string &title);
    
    // freeze panes
    cell_reference get_frozen_panes() const;
    void freeze_panes(cell top_left_cell);
    void freeze_panes(const std::string &top_left_coordinate);
    void unfreeze_panes();
    bool has_frozen_panes() const;
    
    // container
    cell get_cell(const cell_reference &reference);
    range get_range(const range_reference &reference);
    range get_named_range(const std::string &name);
    range rows() const;
    range columns() const;
    std::list<cell> get_cell_collection();
    named_range create_named_range(const std::string &name, const range_reference &reference);
    
    // extents
    int get_highest_row() const;
    int get_highest_column() const;
    range_reference calculate_dimension() const;
    
    // relationships
    relationship create_relationship(const std::string &relationship_type, const std::string &target_uri);
    std::vector<relationship> get_relationships();
    
    // charts
    //void add_chart(chart chart);
    
    // cell merge
    void merge_cells(const range_reference &reference);
    void unmerge_cells(const range_reference &reference);
    std::vector<range_reference> get_merged_ranges() const;
    
    // append
    void append(const std::vector<std::string> &cells);
    void append(const std::unordered_map<std::string, std::string> &cells);
    void append(const std::unordered_map<int, std::string> &cells);

    // operators
    bool operator==(const worksheet &other) const;
    bool operator!=(const worksheet &other) const;
    bool operator==(std::nullptr_t) const;
    bool operator!=(std::nullptr_t) const;
    void operator=(const worksheet &other);
    cell operator[](const cell_reference &reference);
    range operator[](const range_reference &reference);
    range operator[](const std::string &named_range);
    
    // page
    page_setup &get_page_setup();
    margins &get_page_margins();
    
    // auto filter
    range_reference get_auto_filter() const;
    void set_auto_filter(const xlnt::range &range);
    void set_auto_filter(const range_reference &reference);
    void unset_auto_filter();
    bool has_auto_filter() const;
    
private:
    friend class workbook;
    friend class cell;
    
    static worksheet allocate(workbook &owner, const std::string &title);
    static void deallocate(worksheet ws);
    
    worksheet_struct *root_;
};
    
} // namespace xlnt
