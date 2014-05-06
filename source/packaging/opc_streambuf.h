#pragma once

#include <array>
#include <iostream>
#include <opc/opc.h>

namespace xlnt {

class opc_streambuf : public std::streambuf
{
public:
	opc_streambuf(opcContainer *container, const std::string &part_name);
	opc_streambuf(const opc_streambuf &) = delete;
	opc_streambuf &operator=(const opc_streambuf &) = delete;

private:
	int_type underflow();
	int_type uflow();
	int_type pbackfail(int_type ch);
	std::streamsize showmanyc();

	int_type overflow(int_type ch);
	int sync();

	opcContainer *container_;
	std::string part_name_;
	std::array<char, 1024> buffer_;
};

} // namespace xlnt
