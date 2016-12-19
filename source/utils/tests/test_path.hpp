#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <helpers/path_helper.hpp>
#include <helpers/temporary_file.hpp>
#include <xlnt/xlnt.hpp>

class test_path : public CxxTest::TestSuite
{
public:
	void test_exists()
	{
		temporary_file temp;

		if (temp.get_path().exists())
		{
			path_helper::delete_file(temp.get_path());
		}

		TS_ASSERT(!temp.get_path().exists());
		std::ofstream stream(temp.get_path().string());
		TS_ASSERT(temp.get_path().exists());
	}
};
