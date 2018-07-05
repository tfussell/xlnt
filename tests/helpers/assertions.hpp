#pragma once

#include <cmath>
#include <exception>
#include <xlnt/utils/exceptions.hpp>

#define XLNT_STRINGIFYX(x) #x
#define XLNT_STRINGIFY(x) XLNT_STRINGIFYX(x)

#define xlnt_assert(expression)                                               \
    do                                                                        \
    {                                                                         \
        try                                                                   \
        {                                                                     \
            if (expression) break;                                            \
        }                                                                     \
        catch (std::exception & ex)                                           \
        {                                                                     \
            throw ex;                                                         \
        }                                                                     \
        catch (...)                                                           \
        {                                                                     \
        }                                                                     \
        throw xlnt::exception(                                              \
            "assert failed at L:" XLNT_STRINGIFY(__LINE__) "\n" XLNT_STRINGIFY(expression)); \
    } while (false)

#define xlnt_assert_throws_nothing(expression)                                \
    do                                                                        \
    {                                                                         \
        try                                                                   \
        {                                                                     \
            expression;                                                       \
            break;                                                            \
        }                                                                     \
        catch (...)                                                           \
        {                                                                     \
        }                                                                     \
        throw xlnt::exception("assert throws nothing failed at L:" XLNT_STRINGIFY(__LINE__) "\n" XLNT_STRINGIFY(expression)); \
    } while (false)

#define xlnt_assert_throws(expression, exception_type)                        \
    do                                                                        \
    {                                                                         \
        try                                                                   \
        {                                                                     \
            expression;                                                       \
        }                                                                     \
        catch (exception_type)                                                \
        {                                                                     \
            break;                                                            \
        }                                                                     \
        catch (...)                                                           \
        {                                                                     \
        }                                                                     \
        throw xlnt::exception("assert throws failed at L:" XLNT_STRINGIFY(__LINE__) "\n" XLNT_STRINGIFY(expression)); \
    } while (false)

#define xlnt_assert_equals(left, right) xlnt_assert((left) == (right))
#define xlnt_assert_differs(left, right) xlnt_assert((left) != (right))
#define xlnt_assert_delta(left, right, delta) xlnt_assert(std::abs((left) - (right)) <= (delta))
