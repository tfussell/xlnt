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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#include <chrono>
#include <string>
#include <iostream>
#include <sstream> 
#include <iterator>
#include <random>

#include <helpers/timing.hpp>
#include <xlnt/xlnt.hpp>

namespace {

std::size_t random_index(std::size_t max)
{
    static std::random_device rd;
    static std::mt19937 gen(rd());

    std::uniform_int_distribution<> dis(0, static_cast<int>(max - 1));

    return dis(gen);
}

void generate_all_formats(xlnt::workbook &wb, std::vector<xlnt::format>& formats)
{
	const auto vertical_alignments = std::vector<xlnt::vertical_alignment>
	{
		xlnt::vertical_alignment::center,
		xlnt::vertical_alignment::justify,
		xlnt::vertical_alignment::top,
		xlnt::vertical_alignment::bottom
	};

	const auto horizontal_alignments = std::vector<xlnt::horizontal_alignment>
	{
		xlnt::horizontal_alignment::center,
		xlnt::horizontal_alignment::center_continuous,
		xlnt::horizontal_alignment::general,
		xlnt::horizontal_alignment::justify,
		xlnt::horizontal_alignment::left,
		xlnt::horizontal_alignment::right
	};

	const auto font_names = std::vector<std::string>
	{
		"Calibri",
		"Tahoma",
		"Arial",
		"Times New Roman"
	};

	const auto font_sizes = std::vector<double>
	{
		11.,
		13.,
		15.,
		17.,
		19.,
		21.,
		23.,
		25.,
		27.,
		29.,
		31.,
		33.,
		35.
	};

	const auto underline_options = std::vector<xlnt::font::underline_style>
	{
		xlnt::font::underline_style::single,
		xlnt::font::underline_style::none
	};

	for (auto vertical_alignment : vertical_alignments)
	{
		for (auto horizontal_alignment : horizontal_alignments)
		{
			for (auto name : font_names)
			{
				for (auto size : font_sizes)
				{
					for (auto bold : { true, false })
					{
						for (auto underline : underline_options)
						{
							for (auto italic : { true, false })
							{
								auto fmt = wb.create_format();

								xlnt::font f;
								f.name(name);
								f.size(size);
								f.italic(italic);
								f.underline(underline);
								f.bold(bold);
								fmt.font(f);

								xlnt::alignment a;
								a.vertical(vertical_alignment);
								a.horizontal(horizontal_alignment);
								fmt.alignment(a);

								formats.push_back(fmt);
							}
						}
					}
				}
			}
		}
	}
}


xlnt::workbook non_optimized_workbook_formats(int rows_number, int columns_number)
{
	using xlnt::benchmarks::current_time;

	xlnt::workbook wb;
	std::vector<xlnt::format> formats;
	auto start = current_time();

	generate_all_formats(wb, formats);

	auto elapsed = current_time() - start;

	std::cout << "elapsed " << elapsed / 1000.0 << ". generate_all_formats. number of unique formats " << formats.size() << std::endl;

	start = current_time();
	auto worksheet = wb[random_index(wb.sheet_count())];
	auto cells_proceeded = 0;
	for (int row_idx = 1; row_idx <= rows_number; row_idx++)
	{
		for (int col_idx = 1; col_idx <= columns_number; col_idx++)
		{
			auto cell = worksheet.cell(xlnt::cell_reference((xlnt::column_t)col_idx, (xlnt::row_t)row_idx));
			std::ostringstream string_stm;
			string_stm << "Col: " << col_idx << "Row: " << row_idx;
			cell.value(string_stm.str());
			cell.format(formats.at(random_index(formats.size())));
			cells_proceeded++;
		}
	}

	elapsed = current_time() - start;

	std::cout << "elapsed " << elapsed / 1000.0 << ". set values and formats for cells. cells proceeded " << cells_proceeded << std::endl;

	return wb;
}

void to_save_profile(xlnt::workbook &wb, const std::string &f)
{
    using xlnt::benchmarks::current_time;

    auto start = current_time();
    wb.save(f);
    auto elapsed = current_time() - start;

    std::cout << "elapsed " << elapsed / 1000.0 << ". save workbook." << std::endl;
}

void to_load_profile(xlnt::workbook &wb, const std::string &f)
{
	using xlnt::benchmarks::current_time;
	
	auto start = current_time();	
	wb.load(f);
	auto elapsed = current_time() - start;

	std::cout << "elapsed " << elapsed / 1000.0 << ". load workbook." << std::endl;
}

void read_formats_profile(xlnt::workbook &wb, int rows_number, int columns_number)
{
	using xlnt::benchmarks::current_time;

	std::vector<std::string> values;
	std::vector<xlnt::format> formats;
	auto start = current_time();
	auto worksheet = wb[random_index(wb.sheet_count())];
	for (int row_idx = 1; row_idx <= rows_number; row_idx++)
	{
		for (int col_idx = 1; col_idx <= columns_number; col_idx++)
		{
			auto cell = worksheet.cell(xlnt::cell_reference((xlnt::column_t)col_idx, (xlnt::row_t)row_idx));
			values.push_back(cell.value<std::string>());
			formats.push_back(cell.format());
		}
	}

	auto elapsed = current_time() - start;

	std::cout << "elapsed " << elapsed / 1000.0 << ". read values and formats for cells. values count " << values.size() 
		<< ". formats count " << formats.size() << std::endl;
}

} // namespace

int main(int argc, char * argv[])
{
    int rows_number = 1000;
	int columns_number = 10;

	try 
	{
		if (argc > 1)
			rows_number = std::stoi(argv[1]);

		if (argc > 2)
			columns_number = std::stoi(argv[2]);

		std::cout << "started. number of rows " << rows_number << ", number of columns " << columns_number << std::endl;
		auto wb = non_optimized_workbook_formats(rows_number, columns_number);
		auto f = "temp-formats.xlsx";
		to_save_profile(wb, f);

		xlnt::workbook load_formats_wb;
		to_load_profile(load_formats_wb, f);
		read_formats_profile(load_formats_wb, rows_number, columns_number);
	}
	catch(std::exception& ex)
	{
		std::cout << "failed. " << ex.what() << std::endl;
	}

    return 0;
}
