// Copyright (C) 2017 Thomas Fussell
// Copyright (C) 2013 Tomas Kislan
// Copyright (C) 2013 Adam Rudd
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.

#include <array>
#include <detail/cryptography/base64.hpp>

namespace xlnt {
namespace detail {

std::string encode_base64(const std::vector<std::uint8_t> &input)
{
    auto encoded_length = (input.size() + 2 - ((input.size() + 2) % 3)) / 3 * 4;
    auto output = std::string(encoded_length, '\0');
    auto input_iterator = input.begin();
    auto output_iterator = output.begin();
    auto i = std::size_t(0);
    auto j = std::size_t(0);
    std::array<std::uint8_t, 3> a3{{0}};
    std::array<std::uint8_t, 4> a4{{0}};

    const std::string alphabet("ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        "abcdefghijklmnopqrstuvwxyz"
        "0123456789+/");

    while (input_iterator != input.end())
    {
        a3[i++] = *input_iterator++;

        if (i == 3)
        {
            a4[0] = (a3[0] & 0xfc) >> 2;
            a4[1] = static_cast<std::uint8_t>(((a3[0] & 0x03) << 4)
                + ((a3[1] & 0xf0) >> 4));
            a4[2] = static_cast<std::uint8_t>(((a3[1] & 0x0f) << 2)
                + ((a3[2] & 0xc0) >> 6));
            a4[3] = (a3[2] & 0x3f);

            for (i = 0; i < 4; i++)
            {
                *output_iterator++ = alphabet[a4[i]];
            }

            i = 0;
        }
    }

    if (i != 0)
    {
        for (j = i; j < 3; j++)
        {
            a3[j] = 0;
        }

        a4[0] = (a3[0] & 0xfc) >> 2;
        a4[1] = static_cast<std::uint8_t>(((a3[0] & 0x03) << 4)
            + ((a3[1] & 0xf0) >> 4));
        a4[2] = static_cast<std::uint8_t>(((a3[1] & 0x0f) << 2)
            + ((a3[2] & 0xc0) >> 6));
        a4[3] = (a3[2] & 0x3f);

        for (j = 0; j < i + 1; j++)
        {
            *output_iterator++ = alphabet[a4[j]];
        }

        while (i++ < 3)
        {
            *output_iterator++ = '=';
        }
    }

    return output;
}

std::vector<std::uint8_t> decode_base64(const std::string &input)
{
    std::size_t padding_count = 0;
    auto in_end = input.data() + input.size();

    while (*--in_end == '=')
    {
        ++padding_count;
    }

    auto decoded_length = ((6 * input.size()) / 8) - padding_count;
    auto output = std::vector<std::uint8_t>(decoded_length);
    auto input_iterator = input.begin();
    auto output_iterator = output.begin();
    auto i = std::size_t(0);
    auto j = std::size_t(0);
    std::array<std::uint8_t, 3> a3{{0}};
    std::array<std::uint8_t, 4> a4{{0}};

    auto b64_lookup = [](std::uint8_t c) -> std::uint8_t
    {
        if(c >='A' && c <='Z') return c - 'A';
        if(c >='a' && c <='z') return c - 71;
        if(c >='0' && c <='9') return c + 4;
        if(c == '+') return 62;
        if(c == '/') return 63;
        return 255;
    };

    while (input_iterator != input.end())
    {
        if (*input_iterator == '=')
        {
            break;
        }

        a4[i++] = static_cast<std::uint8_t>(*(input_iterator++));

        if (i == 4)
        {
            for (i = 0; i < 4; i++)
            {
                a4[i] = b64_lookup(a4[i]);
            }

            a3[0] = static_cast<std::uint8_t>(a4[0] << 2)
                + static_cast<std::uint8_t>((a4[1] & 0x30) >> 4);
            a3[1] = static_cast<std::uint8_t>((a4[1] & 0xf) << 4)
                + static_cast<std::uint8_t>((a4[2] & 0x3c) >> 2);
            a3[2] = static_cast<std::uint8_t>((a4[2] & 0x3) << 6)
                + static_cast<std::uint8_t>(a4[3]);

            for (i = 0; i < 3; i++)
            {
                *output_iterator++ = a3[i];
            }

            i = 0;
        }
    }

    if (i != 0)
    {
        for (j = i; j < 4; j++)
        {
            a4[j] = 0;
        }

        for (j = 0; j < 4; j++)
        {
            a4[j] = b64_lookup(a4[j]);
        }

        a3[0] = static_cast<std::uint8_t>(a4[0] << 2)
            + static_cast<std::uint8_t>((a4[1] & 0x30) >> 4);
        a3[1] = static_cast<std::uint8_t>((a4[1] & 0xf) << 4)
            + static_cast<std::uint8_t>((a4[2] & 0x3c) >> 2);
        a3[2] = static_cast<std::uint8_t>((a4[2] & 0x3) << 6)
            + static_cast<std::uint8_t>(a4[3]);

        for (j = 0; j < i - 1; j++)
        {
            *output_iterator++ = a3[j];
        }
    }

    return output;
}

} // namespace detail
} // namespace xlnt
