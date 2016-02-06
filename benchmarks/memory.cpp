#include <cassert>

#ifdef __APPLE__
#include<mach/mach.h>
#endif

#include <xlnt/xlnt.hpp>

#include "../tests/helpers/path_helper.hpp"

int calc_memory_usage()
{
#ifdef __APPLE__
    struct task_basic_info t_info;
    mach_msg_type_number_t t_info_count = TASK_BASIC_INFO_COUNT;

    if (KERN_SUCCESS != task_info(mach_task_self(),
				  TASK_BASIC_INFO, (task_info_t)&t_info, 
				  &t_info_count))
    {
	return 0;
    }

    return t_info.virtual_size;
#endif
    return 0;
}

void test_memory_use()
{
    // Naive test that assumes memory use will never be more than 120 % of
    // that for first 50 rows
    auto current_folder = PathHelper::GetExecutableDirectory();
    auto src = current_folder + "rks/files/very_large.xlsx";

    xlnt::workbook wb;
    wb.load(src);
    auto ws = wb.get_active_sheet();

    int initial_use = 0;
    int n = 0;

    for (auto line : ws.rows())
    {
        if (n % 50 == 0)
        {
            auto use = calc_memory_usage();

            if (initial_use == 0)
	    {
                initial_use = use;
	    }

            assert(use / initial_use < 1.2);
	    std::cout << n << " " << use << std::endl;
	}

        n++;
    }
}

int main()
{
    test_memory_use();
}
