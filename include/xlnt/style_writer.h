#pragma once

#include <string>
#include <unordered_map>
#include <vector>

#include "style.h"

namespace xlnt {

class workbook;

class style_writer
{
 public:
  style_writer(workbook &wb);
  std::unordered_map<std::size_t, std::string> get_style_by_hash() const;
  std::string write_table() const;

 private:
  std::vector<style> get_style_list(const workbook &wb) const;
  std::unordered_map<int, std::string> write_fonts() const;
  std::unordered_map<int, std::string> write_fills() const;
  std::unordered_map<int, std::string> write_borders() const;
  void write_cell_style_xfs();
  void write_cell_xfs();
  void write_cell_style();
  void write_dxfs();
  void write_table_styles();
  void write_number_formats();

  std::vector<style> style_list_;
  workbook &wb_;
};

} // namespace xlnt
