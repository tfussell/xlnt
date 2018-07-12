#pragma once

#include <cstdint>
#include <functional>
#include <iostream>
#include <string>
#include <utility>
#include <vector>

#include <helpers/assertions.hpp>
#include <helpers/path_helper.hpp>
//#include <helpers/temporary_directory.hpp>
//#include <helpers/temporary_file.hpp>
#include <helpers/timing.hpp>
#include <helpers/xml_helper.hpp>

struct test_status
{
    std::size_t tests_run = 0;
    std::size_t tests_failed = 0;
    std::size_t tests_passed = 0;
    std::vector<std::string> failures;
};

std::string build_name(const std::string &pretty, const std::string &method);

#define register_test(test) register_test_internal([this]() { test(); }, build_name(__FUNCTION__, #test));

class test_suite
{
public:
    static test_status go();

protected:
    static void register_test_internal(std::function<void()> t, const std::string &function)
    {
        tests().push_back(std::make_pair(t, function));
    }

private:
    static std::vector<std::pair<std::function<void(void)>, std::string>> &tests();
};
