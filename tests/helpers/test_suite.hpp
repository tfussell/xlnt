#pragma once

#include <cstdint>
#include <functional>
#include <iostream>
#include <string>
#include <utility>
#include <vector>

#include <helpers/assertions.hpp>

struct test_status
{
    std::size_t tests_run = 0;
    std::size_t tests_failed = 0;
    std::size_t tests_passed = 0;
    std::vector<std::string> failures;
};

std::string build_name(const std::string &pretty, const std::string &method)
{
    return pretty.substr(0, pretty.find("::") + 2) + method;
}

#define register_test(test) register_test_internal([this]() { test(); }, build_name(__FUNCTION__, #test));

class test_suite
{
public:
    test_status go()
    {
        test_status status;

        for (auto test : tests)
        {
            try
            {
                test.first();
                std::cout << ".";
                status.tests_passed++;
            }
            catch (...)
            {
                std::cout << "*";
                status.tests_failed++;
                status.failures.push_back(test.second);
            }

            std::cout.flush();
            status.tests_run++;
        }

        return status;
    }

protected:
    void register_test_internal(std::function<void()> t, const std::string &function)
    {
        tests.push_back(std::make_pair(t, function));
    }

private:
    std::vector<std::pair<std::function<void(void)>, std::string>> tests;
};
