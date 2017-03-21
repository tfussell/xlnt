#include <chrono>
#include <iostream>
#include <xlnt/xlnt.hpp>

std::size_t current_time()
{
    return static_cast<std::size_t>(std::chrono::duration<double, std::milli>(std::chrono::system_clock::now().time_since_epoch()).count());
}

// Create a worksheet with variable width rows. Because data must be
// serialised row by row it is often the width of the rows which is most
// important.
void writer(int cols, int rows)
{
    xlnt::workbook wb;
    auto ws = wb.create_sheet();

    for(int index = 0; index < rows; index++)
    {
        if ((index + 1) % (rows / 10) == 0)
        {
            std::string progress = std::string((index + 1) / (1 + rows / 10), '.');
            std::cout << "\r" << progress;
            std::cout.flush();
        }

		for (int i = 0; i < cols; i++)
		{
			ws.cell(xlnt::cell_reference(i + 1, index + 1)).value(i);
		}
    }

    std::cout << std::endl;

    auto filename = "data/benchmark.xlsx";
    wb.save(filename);
}

// Create a timeit call to a function and pass in keyword arguments.
// The function is called twice, once using the standard workbook, then with the optimised one.
// Time from the best of three is taken.
int timer(std::function<void(int, int)> fn, int cols, int rows)
{
    const auto repeat = std::size_t(3);
    auto time = std::numeric_limits<std::size_t>::max();

    std::cout << cols << " cols " << rows << " rows" << std::endl;

    for(int i = 0; i < repeat; i++)
    {
        auto start = current_time();
        fn(cols, rows);
        time = std::min(current_time() - start, time);
    }

    std::cout << time / 1000.0 << std::endl;

    return time;
}

int main()
{
    timer(&writer, 100, 100);
    timer(&writer, 1000, 100);
    timer(&writer, 4000, 100);
    timer(&writer, 8192, 100);
    timer(&writer, 10, 10000);
    timer(&writer, 4000, 1000);

    return 0;
}
