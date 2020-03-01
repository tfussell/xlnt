// A core part of the xlsx serialisation routine is taking doubles from memory and stringifying them
// this has a few requirements
// - expect strings in the form 1234.56 (i.e. no thousands seperator, '.' used for the decimal seperator)
// - outputs up to 15 significant figures (excel only serialises numbers up to 15sf)

#include "benchmark/benchmark.h"
#include <locale>
#include <random>
#include <sstream>

namespace {

// setup a large quantity of random doubles as strings
template <bool Decimal_Locale = true>
class RandomFloats : public benchmark::Fixture
{
    static constexpr size_t Number_of_Elements = 1 << 20;
    static_assert(Number_of_Elements > 1'000'000, "ensure a decent set of random values is generated");

    std::vector<double> inputs;

    size_t index = 0;
    const char *locale_str = nullptr;

public:
    void SetUp(const ::benchmark::State &state)
    {
        if (Decimal_Locale)
        {
            locale_str = setlocale(LC_ALL, "C");
        }
        else
        {
            locale_str = setlocale(LC_ALL, "de-DE");
        }
        std::random_device rd; // obtain a seed for the random number engine
        std::mt19937 gen(rd());
        // doing full range is stupid (<double>::min/max()...), it just ends up generating very large numbers
        // uniform is probably not the best distribution to use here, but it will do for now
        std::uniform_real_distribution<double> dis(-1'000, 1'000);
        // generate a large quantity of doubles to deserialise
        inputs.reserve(Number_of_Elements);
        for (int i = 0; i < Number_of_Elements; ++i)
        {
            double d = dis(gen);
            inputs.push_back(d);
        }
    }

    void TearDown(const ::benchmark::State &state)
    {
        // restore locale
        setlocale(LC_ALL, locale_str);
        // gbench is keeping the fixtures alive somewhere, need to clear the data after use
        inputs = std::vector<double>{};
    }

    double &get_rand()
    {
        return inputs[++index & (Number_of_Elements - 1)];
    }
};

/// Takes in a double and outputs a string form of that number which will
/// serialise and deserialise without loss of precision
std::string serialize_number_to_string(double num)
{
    // more digits and excel won't match
    constexpr int Excel_Digit_Precision = 15; //sf
    std::stringstream ss;
    ss.precision(Excel_Digit_Precision);
    ss << num;
    return ss.str();
}

class number_serialiser
{
    static constexpr int Excel_Digit_Precision = 15; //sf
    std::ostringstream ss;

public:
    explicit number_serialiser()
    {
        ss.precision(Excel_Digit_Precision);
        ss.imbue(std::locale("C"));
    }

    std::string serialise(double d)
    {
        ss.str(""); // reset string buffer
        ss.clear(); // reset any error flags
        ss << d;
        return ss.str();
    }
};

class number_serialiser_mk2
{
    static constexpr int Excel_Digit_Precision = 15; //sf
    bool should_convert_comma;

    void convert_comma(char *buf, int len)
    {
        char *buf_end = buf + len;
        char *decimal = std::find(buf, buf_end, ',');
        if (decimal != buf_end)
        {
            *decimal = '.';
        }
    }

public:
    explicit number_serialiser_mk2()
        : should_convert_comma(std::use_facet<std::numpunct<char>>(std::locale{}).decimal_point() == ',')
    {
    }

    std::string serialise(double d)
    {
        char buf[Excel_Digit_Precision + 1]; // need space for trailing '\0'
        int len = snprintf(buf, sizeof(buf), "%.15g", d);
        if (should_convert_comma)
        {
            convert_comma(buf, len);
        }
        return std::string(buf, len);
    }
};

using RandFloats = RandomFloats<true>;
using RandFloatsComma = RandomFloats<false>;
} // namespace

BENCHMARK_F(RandFloats, string_from_double_sstream)
(benchmark::State &state)
{
    while (state.KeepRunning())
    {
        benchmark::DoNotOptimize(
            serialize_number_to_string(get_rand()));
    }
}

BENCHMARK_F(RandFloats, string_from_double_sstream_cached)
(benchmark::State &state)
{
    number_serialiser ser;
    while (state.KeepRunning())
    {
        benchmark::DoNotOptimize(
            ser.serialise(get_rand()));
    }
}

BENCHMARK_F(RandFloats, string_from_double_snprintf)
(benchmark::State &state)
{
    while (state.KeepRunning())
    {
        char buf[16];
        int len = snprintf(buf, sizeof(buf), "%.15g", get_rand());

        benchmark::DoNotOptimize(
            std::string(buf, len));
    }
}

BENCHMARK_F(RandFloats, string_from_double_snprintf_fixed)
(benchmark::State &state)
{
    number_serialiser_mk2 ser;
    while (state.KeepRunning())
    {
        benchmark::DoNotOptimize(
            ser.serialise(get_rand()));
    }
}

// locale names are different between OS's, and std::from_chars is only complete in MSVC
#ifdef _MSC_VER

#include <charconv>
BENCHMARK_F(RandFloats, string_from_double_std_to_chars)
(benchmark::State &state)
{
    while (state.KeepRunning())
    {
        char buf[16];
        std::to_chars_result result = std::to_chars(buf, buf + std::size(buf), get_rand());

        benchmark::DoNotOptimize(
            std::string(buf, result.ptr));
    }
}

BENCHMARK_F(RandFloatsComma, string_from_double_snprintf_fixed_comma)
(benchmark::State &state)
{
    number_serialiser_mk2 ser;
    while (state.KeepRunning())
    {
        benchmark::DoNotOptimize(
            ser.serialise(get_rand()));
    }
}

#endif