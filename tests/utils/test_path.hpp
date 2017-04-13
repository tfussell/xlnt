#pragma once

#include <iostream>
#include <helpers/test_suite.hpp>

#include <helpers/path_helper.hpp>
#include <helpers/temporary_file.hpp>
#include <xlnt/xlnt.hpp>

class test_path : public test_suite
{
public:
	void test_exists()
	{
		temporary_file temp;

		if (temp.get_path().exists())
		{
			path_helper::delete_file(temp.get_path());
		}

		assert(!temp.get_path().exists());
		std::ofstream stream(temp.get_path().string());
		assert(temp.get_path().exists());
	}
};
