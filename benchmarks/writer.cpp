#include <chrono>
#include <iostream>
#include <xlnt/xlnt.hpp>

int current_time()
{
    return std::chrono::duration<double, std::milli>(std::chrono::system_clock::now().time_since_epoch()).count();
}

// Create a worksheet with variable width rows. Because data must be
// serialised row by row it is often the width of the rows which is most
// important.
void writer(bool optimized, int cols, int rows)
{
    xlnt::workbook wb;
    auto ws = wb.create_sheet();

    std::vector<int> row;

    for(int i = 0; i < cols; i++)
    {
	row.push_back(i);
    }

    for(int index = 0; index < rows; index++)
    {
	if ((index + 1) % (rows / 10) == 0)
	{
	    std::string progress = std::string((index + 1) / (1 + rows / 10), '.');
	    std::cout << "\r" << progress;
	    std::cout.flush();
	}

        ws.append(row);
    }

    std::cout << std::endl;

    auto filename = "data/large.xlsx";
    wb.save(filename);
}

// Create a timeit call to a function and pass in keyword arguments.
// The function is called twice, once using the standard workbook, then with the optimised one.
// Time from the best of three is taken.
std::pair<int, int> timer(std::function<void(bool, int, int)> fn, int cols, int rows)
{
    const int repeat = 3;
    int min_time_standard = std::numeric_limits<int>::max();
    int min_time_optimized = std::numeric_limits<int>::max();

    for(bool opt : {false, true})
    {
	std::cout << cols << " cols " << rows << " rows, Worksheet is " << (opt ? "optimised" : "not optimised") << std::endl;
	auto &time = opt ? min_time_optimized : min_time_standard;

        for(int i = 0; i < repeat; i++)
	{
	    auto start = current_time();
	    fn(opt, cols, rows);
	    time = std::min(current_time() - start, time);
	}
    }

    double ratio = min_time_optimized / static_cast<double>(min_time_standard) * 100;
    std::cout << "Optimised takes " << ratio << "% time" << std::endl;

    return {min_time_standard, min_time_optimized};
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
