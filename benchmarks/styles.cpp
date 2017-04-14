// Copyright (c) 2017 Thomas Fussell
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
#include <iostream>
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

std::vector<xlnt::style> generate_all_styles(xlnt::workbook &wb)
{
    std::vector<xlnt::style> styles;

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

    std::size_t index = 0;

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
                                auto s = wb.create_style(std::to_string(index++));

                                xlnt::font f;
                                f.name(name);
                                f.size(size);
                                f.italic(italic);
                                f.underline(underline);
                                f.bold(bold);
                                s.font(f);

                                xlnt::alignment a;
                                a.vertical(vertical_alignment);
                                a.horizontal(horizontal_alignment);
                                s.alignment(a);

                                styles.push_back(s);
                            }
                        }
                    }
                }
            }
        }
    }

    return styles;
}

xlnt::workbook non_optimized_workbook(int n)
{
    xlnt::workbook wb;
    auto styles = generate_all_styles(wb);

    for(int idx = 1; idx < n; idx++)
    {
        auto worksheet = wb[random_index(wb.sheet_count())];
        auto cell = worksheet.cell(xlnt::cell_reference(1, (xlnt::row_t)idx));
        cell.value(0);
        cell.style(styles.at(random_index(styles.size())));
    }

    return wb;
}

void to_profile(xlnt::workbook &wb, const std::string &f, int n)
{
    using xlnt::benchmarks::current_time;

    auto start = current_time();
    wb.save(f);
    auto elapsed = current_time() - start;

    std::cout << "took " << elapsed / 1000.0 << "s for " << n << " styles" << std::endl;
}

} // namespace

int main()
{
    int n = 10000;
    auto wb = non_optimized_workbook(n);
    std::string f = "temp.xlsx";
    to_profile(wb, f, n);

    return 0;
}
