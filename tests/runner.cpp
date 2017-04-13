#include <cell/test_cell.hpp>
#include <cell/test_index_types.hpp>
#include <cell/test_rich_text.hpp>

#include <styles/test_alignment.hpp>
#include <styles/test_color.hpp>
#include <styles/test_fill.hpp>
#include <styles/test_number_format.hpp>

#include <utils/test_datetime.hpp>
#include <utils/test_path.hpp>
#include <utils/test_tests.hpp>
#include <utils/test_timedelta.hpp>
#include <utils/test_units.hpp>

#include <workbook/test_consume_xlsx.hpp>
#include <workbook/test_named_range.hpp>
#include <workbook/test_produce_xlsx.hpp>
#include <workbook/test_protection.hpp>
#include <workbook/test_round_trip.hpp>
#include <workbook/test_theme.hpp>
#include <workbook/test_vba.hpp>
#include <workbook/test_workbook.hpp>

#include <worksheet/test_datavalidation.hpp>
#include <worksheet/test_page_setup.hpp>
#include <worksheet/test_range.hpp>
#include <worksheet/test_worksheet.hpp>

test_status overall_status;

template<typename T>
void run_tests()
{
    auto status = T{}.go();

    overall_status.tests_run += status.tests_run;
    overall_status.tests_passed += status.tests_passed;
    overall_status.tests_failed += status.tests_failed;

    std::copy(status.failures.begin(), status.failures.end(), std::back_inserter(overall_status.failures));
}

void print_summary()
{
    std::cout << std::endl;

    for (auto failure : overall_status.failures)
    {
	std::cout << failure << " failed" << std::endl;
    }
}

int main()
{
    run_tests<cell_test_suite>();
    run_tests<test_index_types>();
    run_tests<test_worksheet>();

    print_summary();

    return overall_status.tests_failed;
}
