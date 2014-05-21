#pragma once

#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

namespace xlnt {

class drawing;
class named_range;
class range_reference;
class relationship;
class worksheet;
    
enum class optimization
{
    write,
    read,
    none
};

class workbook
{
public:
    //constructors
    workbook(optimization optimization = optimization::none);
    ~workbook();
    
    //prevent copy and assignment
    workbook(const workbook &) = delete;
    const workbook &operator=(const workbook &) = delete;
    
    void read_workbook_settings(const std::string &xml_source);
    
    //getters
    worksheet get_active_sheet();
    bool get_optimized_write() const { return optimized_write_; }
    bool get_guess_types() const { return guess_types_; }
    bool get_data_only() const { return data_only_; }
    
    //create
    worksheet create_sheet();
    worksheet create_sheet(std::size_t index);
    worksheet create_sheet(const std::string &title);
    worksheet create_sheet(std::size_t index, const std::string &title);
    
    //add
    void add_sheet(worksheet worksheet);
    void add_sheet(worksheet worksheet, std::size_t index);
    
    //remove
    void remove_sheet(worksheet worksheet);
    void clear();
    
    //container operations
    worksheet get_sheet_by_name(const std::string &sheet_name);
    
    bool contains(const std::string &key) const;
    
    int get_index(worksheet worksheet);
    
    worksheet operator[](const std::string &name);
    worksheet operator[](int index);
    
    std::vector<worksheet>::iterator begin();
    std::vector<worksheet>::iterator end();
    
    std::vector<worksheet>::const_iterator begin() const;
    std::vector<worksheet>::const_iterator end() const;
    
    std::vector<worksheet>::const_iterator cbegin() const;
    std::vector<worksheet>::const_iterator cend() const;
    
    std::vector<std::string> get_sheet_names() const;
    
    //named ranges
    named_range create_named_range(const std::string &name, worksheet worksheet, const range_reference &reference);
    std::vector<named_range> get_named_ranges();
    void add_named_range(const std::string &name, named_range named_range);
    bool has_named_range(const std::string &name, worksheet ws) const;
    named_range get_named_range(const std::string &name, worksheet ws);
    void remove_named_range(named_range named_range);
    
    //serialization
    void save(const std::string &filename);
    void load(const std::string &filename);
    
    bool operator==(const workbook &rhs) const;
    
    std::unordered_map<std::string, std::pair<std::string, std::string>> get_root_relationships() const;
    std::unordered_map<std::string, std::pair<std::string, std::string>> get_relationships() const;
    
private:
    bool optimized_write_;
    bool optimized_read_;
    bool guess_types_;
    bool data_only_;
    int active_sheet_index_;
    std::vector<worksheet> worksheets_;
    std::unordered_map<std::string, named_range> named_ranges_;
    std::vector<relationship> relationships_;
    std::vector<drawing> drawings_;
};
    
} // namespace xlnt
