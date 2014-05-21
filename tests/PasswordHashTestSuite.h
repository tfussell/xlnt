#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.h>

class PasswordHashTestSuite : public CxxTest::TestSuite
{
public:
    PasswordHashTestSuite()
    {

    }

    void test_hasher()
    {
        TS_ASSERT_EQUALS("CBEB", xlnt::sheet_protection::hash_password("test"));
    }

    void test_sheet_protection()
    {
        xlnt::sheet_protection protection;
        protection.set_password("test");
        TS_ASSERT_EQUALS("CBEB", protection.get_hashed_password());
    }
};
