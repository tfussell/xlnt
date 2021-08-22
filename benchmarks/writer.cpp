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

#include <chrono>
#include <iostream>

#include <helpers/timing.hpp>
#include <xlnt/xlnt.hpp>

namespace {

// Create a worksheet with variable width rows. Because data must be
// serialised row by row it is often the width of the rows which is most
// important.
void writer(int cols, int rows)
{
    xlnt::workbook wb;
    auto ws = wb.create_sheet();

    for(int index = 0; index < rows; index++)
    {
        if (rows >= 10 && (index + 1) % (rows / 10) == 0)
        {
            std::string progress = std::string((index + 1) / (1 + rows / 10), '.');
            std::cout << "\r" << progress;
        }

		for (int i = 0; i < cols; i++)
		{
			ws.cell(xlnt::cell_reference(i + 1, index + 1)).value(i);
		}
    }
    std::cout << '\n';

    auto filename = "benchmark.xlsx";
    wb.save(filename);
}

// Create a timeit call to a function and pass in keyword arguments.
// The function is called twice, once using the standard workbook, then with the optimised one.
// Time from the best of three is taken.
void timer(std::function<void(int, int)> fn, int cols, int rows)
{
    const auto repeat = std::size_t(3);
    std::chrono::duration<double, std::milli> time{};
    std::cout << cols << " cols " << rows << " rows" << std::endl;
    fn(rows, cols); // 1 cold run

    for(int i = 0; i < repeat; i++)
    {
        auto start = std::chrono::high_resolution_clock::now();
        fn(cols, rows);
        time += std::chrono::high_resolution_clock::now() - start;
    }

    std::cout << time.count() / repeat << " ms per iteration" << '\n' << '\n';
}

} // namespace

int main()
{
    timer(&writer, 10000, 1);
    timer(&writer, 1000, 10);
    timer(&writer, 100, 100);
    timer(&writer, 10, 1000);
    timer(&writer, 1, 10000);

    return 0;
}
