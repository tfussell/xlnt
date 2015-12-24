// Copyright (c) 2014-2016 Thomas Fussell
// Copyright (c) 2010-2015 openpyxl
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

#include <memory>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/packaging/relationship.hpp>

namespace xlnt {

class alignment;
class border;
class color;
class document_properties;
class drawing;
class fill;
class font;
class manifest;
class named_range;
class number_format;
class pattern_fill;
class protection;
class range;
class range_reference;
class relationship;
class style;
class theme;
class worksheet;
class zip_file;

enum class encoding;

namespace detail { struct workbook_impl; } // namespace detail

/// <summary>
/// workbook is the container for all other parts of the document.
/// </summary>
class XLNT_CLASS workbook
{
  public:
    class XLNT_CLASS iterator
    {
      public:
        iterator(workbook &wb, std::size_t index);
        iterator(const iterator &);
        iterator &operator=(const iterator &);
        worksheet operator*();
        bool operator==(const iterator &comparand) const;
        bool operator!=(const iterator &comparand) const
        {
            return !(*this == comparand);
        }
        iterator operator++(int);
        iterator &operator++();

      private:
        workbook &wb_;
        std::size_t index_;
    };

    class XLNT_CLASS const_iterator
    {
      public:
        const_iterator(const workbook &wb, std::size_t index);
        const_iterator(const const_iterator &);
        const_iterator &operator=(const const_iterator &);
        const worksheet operator*();
        bool operator==(const const_iterator &comparand) const;
        bool operator!=(const const_iterator &comparand) const
        {
            return !(*this == comparand);
        }
        const_iterator operator++(int);
        const_iterator &operator++();

      private:
        const workbook &wb_;
        std::size_t index_;
    };

    static std::size_t index_from_ws_filename(const std::string &filename);

    // constructors
    workbook();
    workbook(encoding e);

    workbook &operator=(workbook other);
    workbook(workbook &&other);
    workbook(const workbook &other);

    friend void swap(workbook &left, workbook &right);

    encoding get_encoding() const;

    worksheet get_active_sheet();

    bool get_guess_types() const;
    void set_guess_types(bool guess);

    bool get_data_only() const;
    void set_data_only(bool data_only);

    // create
    worksheet create_sheet();
    worksheet create_sheet(std::size_t index);
    worksheet create_sheet(const std::string &title);
    worksheet create_sheet(std::size_t index, const std::string &title);
    worksheet create_sheet(const std::string &title, const relationship &rel);

    // add
    void add_sheet(worksheet worksheet);
    void add_sheet(worksheet worksheet, std::size_t index);

    // remove
    void remove_sheet(worksheet worksheet);
    void clear();

    // container operations
    worksheet get_sheet_by_name(const std::string &sheet_name);
    const worksheet get_sheet_by_name(const std::string &sheet_name) const;
    worksheet get_sheet_by_index(std::size_t index);
    const worksheet get_sheet_by_index(std::size_t index) const;
    bool contains(const std::string &key) const;
    int get_index(worksheet worksheet);

    worksheet operator[](const std::string &name);
    worksheet operator[](std::size_t index);

    iterator begin();
    iterator end();

    const_iterator begin() const;
    const_iterator end() const;

    const_iterator cbegin() const;
    const_iterator cend() const;

    std::vector<std::string> get_sheet_names() const;

    document_properties &get_properties();
    const document_properties &get_properties() const;

    // named ranges
    std::vector<named_range> get_named_ranges() const;
    void create_named_range(const std::string &name, worksheet worksheet, const range_reference &reference);
    void create_named_range(const std::string &name, worksheet worksheet, const std::string &reference_string);
    bool has_named_range(const std::string &name) const;
    range get_named_range(const std::string &name);
    void remove_named_range(const std::string &name);

    // serialization
    bool save(std::vector<unsigned char> &data);
    bool save(const std::string &filename);
    bool load(const std::vector<unsigned char> &data);
    bool load(const std::string &filename);
    bool load(std::istream &stream);
    bool load(zip_file &archive);

    bool operator==(const workbook &rhs) const;

    bool operator!=(const workbook &rhs) const
    {
        return !(*this == rhs);
    }

    bool operator==(std::nullptr_t) const;

    bool operator!=(std::nullptr_t) const
    {
        return !(*this == std::nullptr_t{});
    }

    void create_relationship(const std::string &id, const std::string &target, relationship::type type);
    relationship get_relationship(const std::string &id) const;
    const std::vector<relationship> &get_relationships() const;

    void add_alignment(const alignment &a);
    void add_border(const border &b);
    void add_fill(const fill &f);
    void add_font(const font &f);
    void add_color(const color &c);
    void add_number_format(const number_format &format);
    void add_protection(const protection &p);

    const std::vector<style> &get_styles() const;

    const std::vector<color> &get_colors() const;
    const std::vector<border> &get_borders() const;
    const std::vector<fill> &get_fills() const;
    const std::vector<font> &get_fonts() const;
    const std::vector<number_format> &get_number_formats() const;

    color add_indexed_color(const color &rgb_color);
    color get_indexed_color(const color &indexed_color) const;

    const number_format &get_number_format(std::size_t style_id) const;
    std::size_t set_number_format(const number_format &format, std::size_t style_id);
    const font &get_font(std::size_t style_id) const;
    std::size_t set_font(const font &font_, std::size_t style_id);
    const fill &get_fill(std::size_t style_id) const;
    std::size_t set_fill(const fill &fill_, std::size_t style_id);
    const border &get_border(std::size_t style_id) const;
    std::size_t set_border(const border &border_, std::size_t style_id);
    const alignment &get_alignment(std::size_t style_id) const;
    std::size_t set_alignment(const alignment &alignment_, std::size_t style_id);
    const protection &get_protection(std::size_t style_id) const;
    std::size_t set_protection(const protection &protection_, std::size_t style_id);
    bool get_pivot_button(std::size_t style_id) const;
    bool get_quote_prefix(std::size_t style_id) const;

    void set_code_name(const std::string &code_name);

    bool has_loaded_theme() const;
    const theme &get_loaded_theme() const;

    const style &get_style(std::size_t style_id) const;
    std::size_t add_style(const style &style_);

    manifest &get_manifest();
    const manifest &get_manifest() const;

    const std::vector<relationship> &get_root_relationships() const;

    void add_shared_string(const std::string &shared);
    std::vector<std::string> &get_shared_strings();
    const std::vector<std::string> &get_shared_strings() const;

  private:
    friend class worksheet;
    std::shared_ptr<detail::workbook_impl> d_;
};

} // namespace xlnt
