#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "../xlnt.h"

class PasswordHashTestSuite : public CxxTest::TestSuite
{
public:
    PasswordHashTestSuite()
    {

    }

    void test_hasher()
    {
        TS_ASSERT_EQUALS("CBEB", hash_password("test"));
    }

    void test_sheet_protection()
    {
        protection = SheetProtection();
        protection.password = "test";
        TS_ASSERT_EQUALS("CBEB", protection.password);
    }
};
