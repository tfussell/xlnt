#include "writer/style_writer.hpp"
#include "workbook/workbook.hpp"
#include "worksheet/worksheet.hpp"
#include "worksheet/range.hpp"
#include "cell/cell.hpp"

namespace xlnt {

style_writer::style_writer(xlnt::workbook &wb) : wb_(wb)
{
  
}

std::unordered_map<std::size_t, std::string> style_writer::get_style_by_hash() const
{
    std::unordered_map<std::size_t, std::string> styles;
    for(auto ws : wb_)
    {
        for(auto row : ws.rows())
        {
            for(auto cell : row)
            {
                if(cell.has_style())
                {
                    styles[1] = "style";
                }
            }
        }
    }
    return styles;
}

std::vector<style> style_writer::get_styles() const
{
    std::vector<style> styles;
    
    for(auto ws : wb_)
    {
        for(auto row : ws.rows())
        {
            for(auto cell : row)
            {
                if(cell.has_style())
                {
                    styles.push_back(cell.get_style());
                }
            }
        }
    }
    
    return styles;
}
    
std::string style_writer::write_table() const
{
    return "";
}
    
} // namespace xlnt
