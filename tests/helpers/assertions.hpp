#pragma once

#include <cmath>
#include <exception>

#define assert(expression) do\
    {\
        try { if (expression) break; }\
        catch (...) {}\
        throw std::exception();\
    } while (false)

#define assert_throws_nothing(expression) do\
    {\
	try { expression; break; }\
	catch (...) {}\
	throw std::exception();\
    } while (false)

#define assert_throws(expression, exception_type) do\
    {\
	try { expression; }\
	catch (exception_type) { break; }\
	throw std::exception();\
    } while (false)

#define assert_equals(left, right) assert(left == right)
#define assert_differs(left, right) assert(left != right)
#define assert_delta(left, right, delta) assert(std::abs(left - right) <= delta)
