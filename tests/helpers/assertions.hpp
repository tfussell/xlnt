#pragma once

#include <cmath>
#include <exception>

#define xlnt_assert(expression) do\
    {\
        try { if (expression) break; }\
        catch (...) {}\
        throw xlnt::exception("test failed");\
    } while (false)

#define xlnt_assert_throws_nothing(expression) do\
    {\
	try { expression; break; }\
	catch (...) {}\
	throw xlnt::exception("test failed");\
    } while (false)

#define xlnt_assert_throws(expression, exception_type) do\
    {\
	try { expression; }\
	catch (exception_type) { break; }\
    catch (...) {}\
	throw xlnt::exception("test failed");\
    } while (false)

#define xlnt_assert_equals(left, right) xlnt_assert(left == right)
#define xlnt_assert_differs(left, right) xlnt_assert(left != right)
#define xlnt_assert_delta(left, right, delta) xlnt_assert(std::abs(left - right) <= delta)
