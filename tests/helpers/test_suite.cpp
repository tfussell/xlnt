#include "test_suite.hpp"
#include <iostream>

std::vector<std::pair<std::function<void(void)>, std::string>> &test_suite::tests()
{
    static std::vector<std::pair<std::function<void(void)>, std::string>> all_tests;
    return all_tests;
}

std::string build_name(const std::string &pretty, const std::string &method)
{
    return pretty.substr(0, pretty.find("::") + 2) + method;
}

test_status test_suite::go()
{
    test_status status;

    for (auto test : tests())
    {
        try
        {
            test.first();
            std::cout << '.';
            status.tests_passed++;
        }
        catch (std::exception &ex)
        {
            std::string fail_msg = test.second + " failed with:\n" + std::string(ex.what());
            std::cout << "*\n"
                      << fail_msg << '\n';
            status.tests_failed++;
            status.failures.push_back(fail_msg);
        }
        catch (...)
        {
            std::cout << "*\n"
                      << test.second << " failed\n";
            status.tests_failed++;
            status.failures.push_back(test.second);
        }

        std::cout.flush();
        status.tests_run++;
    }

    return status;
}