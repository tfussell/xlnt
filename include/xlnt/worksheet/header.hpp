#pragma once

#include <cstdint>
#include <string>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

/// <summary>
/// Worksheet header
/// </summary>
class XLNT_CLASS header
{
  public:
    header();
    
    void set_text(const std::string &text);
    
    void set_font_name(const std::string &font_name);
    
    void set_font_size(std::size_t font_size);
    
    void set_font_color(const std::string &font_color);
    
    bool is_default() const;

  private:
    bool default_;
    std::string text_;
    std::string font_name_;
    std::size_t font_size_;
    std::string font_color_;
};

} // namespace xlnt
