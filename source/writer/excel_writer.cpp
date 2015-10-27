#include <unordered_set>

#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/common/zip_file.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/writer/excel_writer.hpp>
#include <xlnt/writer/manifest_writer.hpp>
#include <xlnt/writer/workbook_writer.hpp>
#include <xlnt/writer/worksheet_writer.hpp>

#include "detail/constants.hpp"

namespace {

std::vector<std::string> extract_all_strings(xlnt::workbook &wb)
{
    std::unordered_set<std::string> strings;
    
    for(auto ws : wb)
    {
        for(auto row : ws.rows())
        {
            for(auto cell : row)
            {
                if(cell.get_data_type() == xlnt::cell::type::string)
                {
                    strings.insert(cell.get_value<std::string>());
                }
            }
        }
    }
    
    return std::vector<std::string>(strings.begin(), strings.end());
}

} // namespace

namespace xlnt {

excel_writer::excel_writer(workbook &wb) : wb_(wb)
{
}

void excel_writer::save(const std::string &filename, bool as_template)
{
    zip_file archive;
    write_data(archive, as_template);
    archive.save(filename);
}

void excel_writer::write_data(zip_file &archive, bool as_template)
{
    archive.writestr(constants::ArcRootRels, write_root_rels(wb_));
    archive.writestr(constants::ArcWorkbookRels, write_workbook_rels(wb_));
    archive.writestr(constants::ArcApp, write_properties_app(wb_));
    archive.writestr(constants::ArcCore, write_properties_core(wb_.get_properties()));
    
    if(wb_.has_loaded_theme())
    {
        archive.writestr(constants::ArcTheme, wb_.get_loaded_theme());
    }
    else
    {
        archive.writestr(constants::ArcTheme, write_theme());
    }
    
    archive.writestr(constants::ArcWorkbook, write_workbook(wb_));
    
    auto shared_strings = extract_all_strings(wb_);
    
    write_charts(archive);
    write_images(archive);
    write_shared_strings(archive, shared_strings);
    write_worksheets(archive, shared_strings);
    write_chartsheets(archive);
    write_external_links(archive);
    
    style_writer style_writer_(wb_);
    archive.writestr(constants::ArcStyles, style_writer_.write_table());
    
    auto manifest = write_content_types(wb_, as_template);
    archive.writestr(constants::ArcContentTypes, manifest);
}
    
void excel_writer::write_shared_strings(xlnt::zip_file &archive, const std::vector<std::string> &shared_strings)
{
    archive.writestr(constants::ArcSharedString, ::xlnt::write_shared_strings(shared_strings));
}

void excel_writer::write_images(zip_file &/*archive*/)
{
    
}

void excel_writer::write_charts(zip_file &/*archive*/)
{
    
}

void excel_writer::write_chartsheets(zip_file &/*archive*/)
{
    
}

void excel_writer::write_worksheets(zip_file &archive, const std::vector<std::string> &shared_strings)
{
    std::size_t index = 0;
    
    for(auto ws : wb_)
    {
        for(auto relationship : wb_.get_relationships())
        {
            if(relationship.get_type() == relationship::type::worksheet &&
               workbook::index_from_ws_filename(relationship.get_target_uri()) == index)
            {
                archive.writestr(relationship.get_target_uri(), write_worksheet(ws, shared_strings));
                break;
            }
        }
        
        index++;
    }
}

void excel_writer::write_external_links(zip_file &/*archive*/)
{
    
}

bool save_workbook(workbook &wb, const std::string &filename, bool as_template)
{
    excel_writer writer(wb);
    writer.save(filename, as_template);
    return true;
}

std::vector<std::uint8_t> save_virtual_workbook(xlnt::workbook &wb, bool as_template)
{
    zip_file archive;
    excel_writer writer(wb);
    writer.write_data(archive, as_template);
    std::vector<std::uint8_t> buffer;
    archive.save(buffer);
    
    return buffer;
}

} // namespace xlnt
