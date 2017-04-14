/* 
 * SHA-512 hash in C
 * 
 * Copyright (c) 2016 Project Nayuki
 * https://www.nayuki.io/page/fast-sha2-hashes-in-x86-assembly
 * 
 * (MIT License)
 * Permission is hereby granted, free of charge, to any person obtaining a copy of
 * this software and associated documentation files (the "Software"), to deal in
 * the Software without restriction, including without limitation the rights to
 * use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
 * the Software, and to permit persons to whom the Software is furnished to do so,
 * subject to the following conditions:
 * - The above copyright notice and this permission notice shall be included in
 *   all copies or substantial portions of the Software.
 * - The Software is provided "as is", without warranty of any kind, express or
 *   implied, including but not limited to the warranties of merchantability,
 *   fitness for a particular purpose and noninfringement. In no event shall the
 *   authors or copyright holders be liable for any claim, damages or other
 *   liability, whether in an action of contract, tort or otherwise, arising from,
 *   out of or in connection with the Software or the use or other dealings in the
 *   Software.
 */

#include <stddef.h>
#include <stdint.h>
#include <string.h>


void sha512_compress(uint64_t state[8], const uint8_t block[128]) {
	#define ROTR64(x, n)  (((0U + (x)) << (64 - (n))) | ((x) >> (n)))  // Assumes that x is uint64_t and 0 < n < 64
	
	#define LOADSCHEDULE(i)  \
		schedule[i] = (uint64_t)block[i * 8 + 0] << 56  \
		            | (uint64_t)block[i * 8 + 1] << 48  \
		            | (uint64_t)block[i * 8 + 2] << 40  \
		            | (uint64_t)block[i * 8 + 3] << 32  \
		            | (uint64_t)block[i * 8 + 4] << 24  \
		            | (uint64_t)block[i * 8 + 5] << 16  \
		            | (uint64_t)block[i * 8 + 6] <<  8  \
		            | (uint64_t)block[i * 8 + 7] <<  0;
	
	#define SCHEDULE(i)  \
		schedule[i] = 0U + schedule[i - 16] + schedule[i - 7]  \
			+ (ROTR64(schedule[i - 15], 1) ^ ROTR64(schedule[i - 15], 8) ^ (schedule[i - 15] >> 7))  \
			+ (ROTR64(schedule[i - 2], 19) ^ ROTR64(schedule[i - 2], 61) ^ (schedule[i - 2] >> 6));
	
	#define ROUND(a, b, c, d, e, f, g, h, i, k) \
		h = 0U + h + (ROTR64(e, 14) ^ ROTR64(e, 18) ^ ROTR64(e, 41)) + (g ^ (e & (f ^ g))) + UINT64_C(k) + schedule[i];  \
		d = 0U + d + h;  \
		h = 0U + h + (ROTR64(a, 28) ^ ROTR64(a, 34) ^ ROTR64(a, 39)) + ((a & (b | c)) | (b & c));
	
	uint64_t schedule[80];
	LOADSCHEDULE( 0)
	LOADSCHEDULE( 1)
	LOADSCHEDULE( 2)
	LOADSCHEDULE( 3)
	LOADSCHEDULE( 4)
	LOADSCHEDULE( 5)
	LOADSCHEDULE( 6)
	LOADSCHEDULE( 7)
	LOADSCHEDULE( 8)
	LOADSCHEDULE( 9)
	LOADSCHEDULE(10)
	LOADSCHEDULE(11)
	LOADSCHEDULE(12)
	LOADSCHEDULE(13)
	LOADSCHEDULE(14)
	LOADSCHEDULE(15)
	SCHEDULE(16)
	SCHEDULE(17)
	SCHEDULE(18)
	SCHEDULE(19)
	SCHEDULE(20)
	SCHEDULE(21)
	SCHEDULE(22)
	SCHEDULE(23)
	SCHEDULE(24)
	SCHEDULE(25)
	SCHEDULE(26)
	SCHEDULE(27)
	SCHEDULE(28)
	SCHEDULE(29)
	SCHEDULE(30)
	SCHEDULE(31)
	SCHEDULE(32)
	SCHEDULE(33)
	SCHEDULE(34)
	SCHEDULE(35)
	SCHEDULE(36)
	SCHEDULE(37)
	SCHEDULE(38)
	SCHEDULE(39)
	SCHEDULE(40)
	SCHEDULE(41)
	SCHEDULE(42)
	SCHEDULE(43)
	SCHEDULE(44)
	SCHEDULE(45)
	SCHEDULE(46)
	SCHEDULE(47)
	SCHEDULE(48)
	SCHEDULE(49)
	SCHEDULE(50)
	SCHEDULE(51)
	SCHEDULE(52)
	SCHEDULE(53)
	SCHEDULE(54)
	SCHEDULE(55)
	SCHEDULE(56)
	SCHEDULE(57)
	SCHEDULE(58)
	SCHEDULE(59)
	SCHEDULE(60)
	SCHEDULE(61)
	SCHEDULE(62)
	SCHEDULE(63)
	SCHEDULE(64)
	SCHEDULE(65)
	SCHEDULE(66)
	SCHEDULE(67)
	SCHEDULE(68)
	SCHEDULE(69)
	SCHEDULE(70)
	SCHEDULE(71)
	SCHEDULE(72)
	SCHEDULE(73)
	SCHEDULE(74)
	SCHEDULE(75)
	SCHEDULE(76)
	SCHEDULE(77)
	SCHEDULE(78)
	SCHEDULE(79)
	
	uint64_t a = state[0];
	uint64_t b = state[1];
	uint64_t c = state[2];
	uint64_t d = state[3];
	uint64_t e = state[4];
	uint64_t f = state[5];
	uint64_t g = state[6];
	uint64_t h = state[7];
	ROUND(a, b, c, d, e, f, g, h,  0, 0x428A2F98D728AE22)
	ROUND(h, a, b, c, d, e, f, g,  1, 0x7137449123EF65CD)
	ROUND(g, h, a, b, c, d, e, f,  2, 0xB5C0FBCFEC4D3B2F)
	ROUND(f, g, h, a, b, c, d, e,  3, 0xE9B5DBA58189DBBC)
	ROUND(e, f, g, h, a, b, c, d,  4, 0x3956C25BF348B538)
	ROUND(d, e, f, g, h, a, b, c,  5, 0x59F111F1B605D019)
	ROUND(c, d, e, f, g, h, a, b,  6, 0x923F82A4AF194F9B)
	ROUND(b, c, d, e, f, g, h, a,  7, 0xAB1C5ED5DA6D8118)
	ROUND(a, b, c, d, e, f, g, h,  8, 0xD807AA98A3030242)
	ROUND(h, a, b, c, d, e, f, g,  9, 0x12835B0145706FBE)
	ROUND(g, h, a, b, c, d, e, f, 10, 0x243185BE4EE4B28C)
	ROUND(f, g, h, a, b, c, d, e, 11, 0x550C7DC3D5FFB4E2)
	ROUND(e, f, g, h, a, b, c, d, 12, 0x72BE5D74F27B896F)
	ROUND(d, e, f, g, h, a, b, c, 13, 0x80DEB1FE3B1696B1)
	ROUND(c, d, e, f, g, h, a, b, 14, 0x9BDC06A725C71235)
	ROUND(b, c, d, e, f, g, h, a, 15, 0xC19BF174CF692694)
	ROUND(a, b, c, d, e, f, g, h, 16, 0xE49B69C19EF14AD2)
	ROUND(h, a, b, c, d, e, f, g, 17, 0xEFBE4786384F25E3)
	ROUND(g, h, a, b, c, d, e, f, 18, 0x0FC19DC68B8CD5B5)
	ROUND(f, g, h, a, b, c, d, e, 19, 0x240CA1CC77AC9C65)
	ROUND(e, f, g, h, a, b, c, d, 20, 0x2DE92C6F592B0275)
	ROUND(d, e, f, g, h, a, b, c, 21, 0x4A7484AA6EA6E483)
	ROUND(c, d, e, f, g, h, a, b, 22, 0x5CB0A9DCBD41FBD4)
	ROUND(b, c, d, e, f, g, h, a, 23, 0x76F988DA831153B5)
	ROUND(a, b, c, d, e, f, g, h, 24, 0x983E5152EE66DFAB)
	ROUND(h, a, b, c, d, e, f, g, 25, 0xA831C66D2DB43210)
	ROUND(g, h, a, b, c, d, e, f, 26, 0xB00327C898FB213F)
	ROUND(f, g, h, a, b, c, d, e, 27, 0xBF597FC7BEEF0EE4)
	ROUND(e, f, g, h, a, b, c, d, 28, 0xC6E00BF33DA88FC2)
	ROUND(d, e, f, g, h, a, b, c, 29, 0xD5A79147930AA725)
	ROUND(c, d, e, f, g, h, a, b, 30, 0x06CA6351E003826F)
	ROUND(b, c, d, e, f, g, h, a, 31, 0x142929670A0E6E70)
	ROUND(a, b, c, d, e, f, g, h, 32, 0x27B70A8546D22FFC)
	ROUND(h, a, b, c, d, e, f, g, 33, 0x2E1B21385C26C926)
	ROUND(g, h, a, b, c, d, e, f, 34, 0x4D2C6DFC5AC42AED)
	ROUND(f, g, h, a, b, c, d, e, 35, 0x53380D139D95B3DF)
	ROUND(e, f, g, h, a, b, c, d, 36, 0x650A73548BAF63DE)
	ROUND(d, e, f, g, h, a, b, c, 37, 0x766A0ABB3C77B2A8)
	ROUND(c, d, e, f, g, h, a, b, 38, 0x81C2C92E47EDAEE6)
	ROUND(b, c, d, e, f, g, h, a, 39, 0x92722C851482353B)
	ROUND(a, b, c, d, e, f, g, h, 40, 0xA2BFE8A14CF10364)
	ROUND(h, a, b, c, d, e, f, g, 41, 0xA81A664BBC423001)
	ROUND(g, h, a, b, c, d, e, f, 42, 0xC24B8B70D0F89791)
	ROUND(f, g, h, a, b, c, d, e, 43, 0xC76C51A30654BE30)
	ROUND(e, f, g, h, a, b, c, d, 44, 0xD192E819D6EF5218)
	ROUND(d, e, f, g, h, a, b, c, 45, 0xD69906245565A910)
	ROUND(c, d, e, f, g, h, a, b, 46, 0xF40E35855771202A)
	ROUND(b, c, d, e, f, g, h, a, 47, 0x106AA07032BBD1B8)
	ROUND(a, b, c, d, e, f, g, h, 48, 0x19A4C116B8D2D0C8)
	ROUND(h, a, b, c, d, e, f, g, 49, 0x1E376C085141AB53)
	ROUND(g, h, a, b, c, d, e, f, 50, 0x2748774CDF8EEB99)
	ROUND(f, g, h, a, b, c, d, e, 51, 0x34B0BCB5E19B48A8)
	ROUND(e, f, g, h, a, b, c, d, 52, 0x391C0CB3C5C95A63)
	ROUND(d, e, f, g, h, a, b, c, 53, 0x4ED8AA4AE3418ACB)
	ROUND(c, d, e, f, g, h, a, b, 54, 0x5B9CCA4F7763E373)
	ROUND(b, c, d, e, f, g, h, a, 55, 0x682E6FF3D6B2B8A3)
	ROUND(a, b, c, d, e, f, g, h, 56, 0x748F82EE5DEFB2FC)
	ROUND(h, a, b, c, d, e, f, g, 57, 0x78A5636F43172F60)
	ROUND(g, h, a, b, c, d, e, f, 58, 0x84C87814A1F0AB72)
	ROUND(f, g, h, a, b, c, d, e, 59, 0x8CC702081A6439EC)
	ROUND(e, f, g, h, a, b, c, d, 60, 0x90BEFFFA23631E28)
	ROUND(d, e, f, g, h, a, b, c, 61, 0xA4506CEBDE82BDE9)
	ROUND(c, d, e, f, g, h, a, b, 62, 0xBEF9A3F7B2C67915)
	ROUND(b, c, d, e, f, g, h, a, 63, 0xC67178F2E372532B)
	ROUND(a, b, c, d, e, f, g, h, 64, 0xCA273ECEEA26619C)
	ROUND(h, a, b, c, d, e, f, g, 65, 0xD186B8C721C0C207)
	ROUND(g, h, a, b, c, d, e, f, 66, 0xEADA7DD6CDE0EB1E)
	ROUND(f, g, h, a, b, c, d, e, 67, 0xF57D4F7FEE6ED178)
	ROUND(e, f, g, h, a, b, c, d, 68, 0x06F067AA72176FBA)
	ROUND(d, e, f, g, h, a, b, c, 69, 0x0A637DC5A2C898A6)
	ROUND(c, d, e, f, g, h, a, b, 70, 0x113F9804BEF90DAE)
	ROUND(b, c, d, e, f, g, h, a, 71, 0x1B710B35131C471B)
	ROUND(a, b, c, d, e, f, g, h, 72, 0x28DB77F523047D84)
	ROUND(h, a, b, c, d, e, f, g, 73, 0x32CAAB7B40C72493)
	ROUND(g, h, a, b, c, d, e, f, 74, 0x3C9EBE0A15C9BEBC)
	ROUND(f, g, h, a, b, c, d, e, 75, 0x431D67C49C100D4C)
	ROUND(e, f, g, h, a, b, c, d, 76, 0x4CC5D4BECB3E42B6)
	ROUND(d, e, f, g, h, a, b, c, 77, 0x597F299CFC657E2A)
	ROUND(c, d, e, f, g, h, a, b, 78, 0x5FCB6FAB3AD6FAEC)
	ROUND(b, c, d, e, f, g, h, a, 79, 0x6C44198C4A475817)
	state[0] = 0U + state[0] + a;
	state[1] = 0U + state[1] + b;
	state[2] = 0U + state[2] + c;
	state[3] = 0U + state[3] + d;
	state[4] = 0U + state[4] + e;
	state[5] = 0U + state[5] + f;
	state[6] = 0U + state[6] + g;
	state[7] = 0U + state[7] + h;
}


void sha512_hash(const uint8_t *message, size_t len, uint64_t hash[8]) {
    hash[0] = UINT64_C(0x6A09E667F3BCC908);
    hash[1] = UINT64_C(0xBB67AE8584CAA73B);
    hash[2] = UINT64_C(0x3C6EF372FE94F82B);
    hash[3] = UINT64_C(0xA54FF53A5F1D36F1);
    hash[4] = UINT64_C(0x510E527FADE682D1);
    hash[5] = UINT64_C(0x9B05688C2B3E6C1F);
    hash[6] = UINT64_C(0x1F83D9ABFB41BD6B);
    hash[7] = UINT64_C(0x5BE0CD19137E2179);

#define BLOCK_SIZE 128  // In bytes
#define LENGTH_SIZE 16  // In bytes

    size_t off;
    for (off = 0; len - off >= BLOCK_SIZE; off += BLOCK_SIZE)
        sha512_compress(hash, &message[off]);

    uint8_t block[BLOCK_SIZE] = { 0 };
    size_t rem = len - off;
    memcpy(block, &message[off], rem);

    block[rem] = 0x80;
    rem++;
    if (BLOCK_SIZE - rem < LENGTH_SIZE) {
        sha512_compress(hash, block);
        memset(block, 0, sizeof(block));
    }

    block[BLOCK_SIZE - 1] = (uint8_t)((len & 0x1FU) << 3);
    len >>= 5;
    for (int i = 1; i < LENGTH_SIZE; i++, len >>= 8)
        block[BLOCK_SIZE - 1 - i] = (uint8_t)(len & 0xFFU);
    sha512_compress(hash, block);
}
