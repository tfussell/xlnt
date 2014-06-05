// Copyright (c) 2014 Thomas Fussell
// Copyright (c) 2010-2014 openpyxl
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file
#pragma once

#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

namespace xlnt {

class drawing;
class range;
class range_reference;
class relationship;
class worksheet;

namespace detail {    
struct workbook_impl;
} // namespace detail
    
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
    
    workbook &operator=(const workbook &) = delete;
    workbook(const workbook &) = delete;
    
    void read_workbook_settings(const std::string &xml_source);
    
    //getters
    worksheet get_active_sheet();
    bool get_optimized_write() const;
    bool get_guess_types() const;
    bool get_data_only() const;
    
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
    const worksheet get_sheet_by_name(const std::string &sheet_name) const;
    worksheet get_sheet_by_index(std::size_t index);
    const worksheet get_sheet_by_index(std::size_t index) const;
    bool contains(const std::string &key) const;
    int get_index(worksheet worksheet);
    
    worksheet operator[](const std::string &name);
    worksheet operator[](int index);
    
    class iterator
    {
    public:
        iterator(workbook &wb, std::size_t index);
        worksheet operator*();
        bool operator==(const iterator &comparand) const;
        bool operator!=(const iterator &comparand) const { return !(*this == comparand); }
        iterator operator++(int);
        iterator &operator++();
        
    private:
        workbook &wb_;
        std::size_t index_;
    };
    
    iterator begin();
    iterator end();
    
    class const_iterator
    {
    public:
        const_iterator(const workbook &wb, std::size_t index);
        const worksheet operator*();
        bool operator==(const const_iterator &comparand) const;
        bool operator!=(const const_iterator &comparand) const { return !(*this == comparand); }
        const_iterator operator++(int);
        const_iterator &operator++();
        
    private:
        const workbook &wb_;
        std::size_t index_;
    };
    
    const_iterator begin() const { return cbegin(); }
    const_iterator end() const { return cend(); }
    
    const_iterator cbegin() const;
    const_iterator cend() const;
    
    std::vector<std::string> get_sheet_names() const;
    
    //named ranges
    void create_named_range(const std::string &name, worksheet worksheet, const range_reference &reference);
    bool has_named_range(const std::string &name) const;
    range get_named_range(const std::string &name);
    void remove_named_range(const std::string &name);
    
    //serialization
    bool save(std::vector<unsigned char> &data);
    bool save(const std::string &filename);
    bool load(const std::vector<unsigned char> &data);
    bool load(const std::string &filename);
    
    bool operator==(const workbook &rhs) const;
    
    std::unordered_map<std::string, std::pair<std::string, std::string>> get_root_relationships() const;
    std::unordered_map<std::string, std::pair<std::string, std::string>> get_relationships() const;
    
private:
    friend class worksheet;
    detail::workbook_impl *d_;
};
    
} // namespace xlnt
