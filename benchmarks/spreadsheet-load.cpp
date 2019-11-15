#include <xlnt/xlnt.hpp>
#include <chrono>
#include <helpers/path_helper.hpp>

namespace {
using milliseconds_d = std::chrono::duration<double, std::milli>;

void run_test(const xlnt::path &file, int runs = 10)
{
    std::cout << file.string() << "\n\n";

    xlnt::workbook wb;
    std::vector<std::chrono::steady_clock::duration> test_timings;

    for (int i = 0; i < runs; ++i)
    {
        auto start = std::chrono::steady_clock::now();
        wb.load(file);

        auto end = std::chrono::steady_clock::now();
        wb.clear();
        test_timings.push_back(end - start);

        std::cout << milliseconds_d(test_timings.back()).count() << " ms\n";
    }
}
} // namespace

int main()
{
    run_test(path_helper::benchmark_file("large.xlsx"));
    run_test(path_helper::benchmark_file("very_large.xlsx"));
}