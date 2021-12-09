// Copyright (c) 2017-2018 Thomas Fussell
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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#include <xlnt/xlnt.hpp>
#include "helpers/path_helper.hpp"

// Readme
// from https://tfussell.gitbooks.io/xlnt/content/ and https://github.com/tfussell/xlnt/blob/master/README.md
void sample_readme_example1()
{
    xlnt::workbook wb;
    xlnt::worksheet ws = wb.active_sheet();

    ws.cell("A1").value(5);
    ws.cell("B2").value("string data");
    ws.cell("C3").formula("=RAND()");

    ws.merge_cells("C3:C4");
    ws.freeze_panes("B2");

    wb.save("sample.xlsx");
}

// Simple - reading from an existing xlsx spread sheet.
// from https://tfussell.gitbooks.io/xlnt/content/docs/introduction/Examples.html
void sample_read_and_print_example()
{
    xlnt::workbook wb;
    wb.load(path_helper::sample_file("documentation-print.xlsx")); // modified to use the test data directory
    auto ws = wb.active_sheet();
    std::clog << "Processing spread sheet" << std::endl;
    for (auto row : ws.rows(false))
    {
        for (auto cell : row)
        {
            std::clog << cell.to_string() << std::endl;
        }
    }
    std::clog << "Processing complete" << std::endl;
}

// Simple - storing a spread sheet in a 2 dimensional C++ Vector for further processing
// from https://tfussell.gitbooks.io/xlnt/content/docs/introduction/Examples.html
void sample_read_into_vector_example()
{
    xlnt::workbook wb;
    wb.load(path_helper::sample_file("documentation-print.xlsx")); // modified to use the test data directory
    auto ws = wb.active_sheet();
    std::clog << "Processing spread sheet" << std::endl;
    std::clog << "Creating a single vector which stores the whole spread sheet" << std::endl;
    std::vector<std::vector<std::string>> theWholeSpreadSheet;
    for (auto row : ws.rows(false))
    {
        std::clog << "Creating a fresh vector for just this row in the spread sheet" << std::endl;
        std::vector<std::string> aSingleRow;
        for (auto cell : row)
        {
            std::clog << "Adding this cell to the row" << std::endl;
            aSingleRow.push_back(cell.to_string());
        }
        std::clog << "Adding this entire row to the vector which stores the whole spread sheet" << std::endl;
        theWholeSpreadSheet.push_back(aSingleRow);
    }
    std::clog << "Processing complete" << std::endl;
    std::clog << "Reading the vector and printing output to the screen" << std::endl;
    for (int rowInt = 0; rowInt < theWholeSpreadSheet.size(); rowInt++)
    {
        for (int colInt = 0; colInt < theWholeSpreadSheet.at(rowInt).size(); colInt++)
        {
            std::cout << theWholeSpreadSheet.at(rowInt).at(colInt) << std::endl;
        }
    }
}

// Simple - writing values to a new xlsx spread sheet.
// from https://tfussell.gitbooks.io/xlnt/content/docs/introduction/Examples.html
void sample_write_sheet_to_file_example()
{
    // Creating a 2 dimensional vector which we will write values to
    std::vector<std::vector<std::string>> wholeWorksheet;
    //Looping through each row (100 rows as per the second argument in the for loop)
    for (int outer = 0; outer < 100; outer++)
    {
        //Creating a fresh vector for a fresh row
        std::vector<std::string> singleRow;
        //Looping through each of the columns (100 as per the second argument in the for loop) in this particular row
        for (int inner = 0; inner < 100; inner++)
        {
            //Adding a single value in each cell of the row
            std::string val = std::to_string(inner + 1);
            singleRow.push_back(val);
        }
        //Adding the single row to the 2 dimensional vector
        wholeWorksheet.push_back(singleRow);
        std::clog << "Writing to row " << outer << " in the vector " << std::endl;
    }
    //Writing to the spread sheet
    //Creating the output workbook
    std::clog << "Creating workbook" << std::endl;
    xlnt::workbook wbOut;
    //Setting the destination output file name
    std::string dest_filename = "output.xlsx";
    //Creating the output worksheet
    xlnt::worksheet wsOut = wbOut.active_sheet();
    //Giving the output worksheet a title/name
    wsOut.title("data");
    //We will now be looping through the 2 dimensional vector which we created above
    //In this case we have two iterators one for the outer loop (row) and one for the inner loop (column)
    std::clog << "Looping through vector and writing to spread sheet" << std::endl;
    for (int fOut = 0; fOut < wholeWorksheet.size(); fOut++)
    {
        std::clog << "Row" << fOut << std::endl;
        for (int fIn = 0; fIn < wholeWorksheet.at(fOut).size(); fIn++)
        {
            //Take notice of the difference between accessing the vector and accessing the work sheet
            //As you may already know Excel spread sheets start at row 1 and column 1 (not row 0 and column 0 like you would expect from a C++ vector)
            //In short the xlnt cell reference starts at column 1 row 1 (hence the + 1s below) and the vector reference starts at row 0 and column 0
            wsOut.cell(xlnt::cell_reference(fIn + 1, fOut + 1)).value(wholeWorksheet.at(fOut).at(fIn));
            //Further clarification to avoid confusion
            //Cell reference arguments are (column number, row number); e.g. cell_reference(fIn + 1, fOut + 1)
            //Vector arguments are (row number, column number); e.g. wholeWorksheet.at(fOut).at(fIn)
        }
    }
    std::clog << "Finished writing spread sheet" << std::endl;
    wbOut.save(dest_filename);
}

// Number Formatting
// from https://tfussell.gitbooks.io/xlnt/content/docs/advanced/Formatting.html
void sample_number_formatting_example()
{
    xlnt::workbook wb;
    auto cell = wb.active_sheet().cell("A1");
    cell.number_format(xlnt::number_format::percentage());
    cell.value(0.513);
    std::cout << '\n'
              << cell.to_string() << std::endl;
}
// Properties
// from https://tfussell.gitbooks.io/xlnt/content/docs/advanced/Properties.html
void sample_properties_example()
{
    xlnt::workbook wb;
    wb.core_property(xlnt::core_property::category, "hors categorie");
    wb.core_property(xlnt::core_property::content_status, "good");
    wb.core_property(xlnt::core_property::created, xlnt::datetime(2017, 1, 15));
    wb.core_property(xlnt::core_property::creator, "me");
    wb.core_property(xlnt::core_property::description, "description");
    wb.core_property(xlnt::core_property::identifier, "id");
    wb.core_property(xlnt::core_property::keywords, {"wow", "such"});
    wb.core_property(xlnt::core_property::language, "Esperanto");
    wb.core_property(xlnt::core_property::last_modified_by, "someone");
    wb.core_property(xlnt::core_property::last_printed, xlnt::datetime(2017, 1, 15));
    wb.core_property(xlnt::core_property::modified, xlnt::datetime(2017, 1, 15));
    wb.core_property(xlnt::core_property::revision, "3");
    wb.core_property(xlnt::core_property::subject, "subject");
    wb.core_property(xlnt::core_property::title, "title");
    wb.core_property(xlnt::core_property::version, "1.0");
    wb.extended_property(xlnt::extended_property::application, "xlnt");
    wb.extended_property(xlnt::extended_property::app_version, "0.9.3");
    wb.extended_property(xlnt::extended_property::characters, 123);
    wb.extended_property(xlnt::extended_property::characters_with_spaces, 124);
    wb.extended_property(xlnt::extended_property::company, "Incorporated Inc.");
    wb.extended_property(xlnt::extended_property::dig_sig, "?");
    wb.extended_property(xlnt::extended_property::doc_security, 0);
    wb.extended_property(xlnt::extended_property::heading_pairs, true);
    wb.extended_property(xlnt::extended_property::hidden_slides, false);
    wb.extended_property(xlnt::extended_property::h_links, 0);
    wb.extended_property(xlnt::extended_property::hyperlink_base, 0);
    wb.extended_property(xlnt::extended_property::hyperlinks_changed, true);
    wb.extended_property(xlnt::extended_property::lines, 42);
    wb.extended_property(xlnt::extended_property::links_up_to_date, false);
    wb.extended_property(xlnt::extended_property::manager, "johnny");
    wb.extended_property(xlnt::extended_property::m_m_clips, "?");
    wb.extended_property(xlnt::extended_property::notes, "note");
    wb.extended_property(xlnt::extended_property::pages, 19);
    wb.extended_property(xlnt::extended_property::paragraphs, 18);
    wb.extended_property(xlnt::extended_property::presentation_format, "format");
    wb.extended_property(xlnt::extended_property::scale_crop, true);
    wb.extended_property(xlnt::extended_property::shared_doc, false);
    wb.extended_property(xlnt::extended_property::slides, 17);
    wb.extended_property(xlnt::extended_property::template_, "template!");
    wb.extended_property(xlnt::extended_property::titles_of_parts, {"title"});
    wb.extended_property(xlnt::extended_property::total_time, 16);
    wb.extended_property(xlnt::extended_property::words, 101);
    wb.custom_property("test", {1, 2, 3});
    wb.custom_property("Editor", "John Smith");
    wb.save("lots_of_properties.xlsx");
}

int main()
{
    sample_readme_example1();
    std::clog << '\n';
    sample_read_and_print_example();
    std::clog << '\n';
    sample_read_into_vector_example();
    std::clog << '\n';
    sample_write_sheet_to_file_example();
    std::clog << '\n';
    sample_number_formatting_example();
    std::clog << '\n';
    sample_properties_example();
    std::clog.flush();
    return 0;
}