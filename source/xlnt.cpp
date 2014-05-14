#include <algorithm>
#include <array>
#include <cassert>
#include <fstream>
#include <iostream>
#include <locale>
#include <sstream>

#include "xlnt.h"
#include "../third-party/pugixml/src/pugixml.hpp"

namespace xlnt {

namespace {

const std::array<unsigned char, 3086> existing_xlsx = {
    0x50, 0x4b, 0x03, 0x04, 0x14, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00,
    0x21, 0x00, 0xf8, 0x17, 0x86, 0x86, 0x7a, 0x01, 0x00, 0x00, 0x10, 0x03,
    0x00, 0x00, 0x10, 0x00, 0x00, 0x00, 0x64, 0x6f, 0x63, 0x50, 0x72, 0x6f,
    0x70, 0x73, 0x2f, 0x61, 0x70, 0x70, 0x2e, 0x78, 0x6d, 0x6c, 0x9d, 0x92,
    0x41, 0x6f, 0xdb, 0x30, 0x0c, 0x85, 0xef, 0x03, 0xf6, 0x1f, 0x0c, 0xdd,
    0x1b, 0x39, 0x5d, 0x31, 0x0c, 0x81, 0xac, 0xa2, 0x48, 0x3b, 0xf4, 0xb0,
    0x61, 0x01, 0x92, 0x76, 0x67, 0x4e, 0xa6, 0x63, 0xa1, 0xb2, 0x24, 0x88,
    0xac, 0x91, 0xec, 0xd7, 0x4f, 0x76, 0x10, 0xd7, 0x59, 0x77, 0xda, 0xed,
    0x91, 0x7c, 0x78, 0xfe, 0x4c, 0x4a, 0xdd, 0x1e, 0x3a, 0x57, 0xf4, 0x98,
    0xc8, 0x06, 0x5f, 0x89, 0xe5, 0xa2, 0x14, 0x05, 0x7a, 0x13, 0x6a, 0xeb,
    0xf7, 0x95, 0x78, 0xda, 0x7d, 0xbd, 0xfa, 0x22, 0x0a, 0x62, 0xf0, 0x35,
    0xb8, 0xe0, 0xb1, 0x12, 0x47, 0x24, 0x71, 0xab, 0x3f, 0x7e, 0x50, 0x9b,
    0x14, 0x22, 0x26, 0xb6, 0x48, 0x45, 0x8e, 0xf0, 0x54, 0x89, 0x96, 0x39,
    0xae, 0xa4, 0x24, 0xd3, 0x62, 0x07, 0xb4, 0xc8, 0x63, 0x9f, 0x27, 0x4d,
    0x48, 0x1d, 0x70, 0x2e, 0xd3, 0x5e, 0x86, 0xa6, 0xb1, 0x06, 0xef, 0x83,
    0x79, 0xed, 0xd0, 0xb3, 0xbc, 0x2e, 0xcb, 0xcf, 0x12, 0x0f, 0x8c, 0xbe,
    0xc6, 0xfa, 0x2a, 0x4e, 0x81, 0xe2, 0x94, 0xb8, 0xea, 0xf9, 0x7f, 0x43,
    0xeb, 0x60, 0x06, 0x3e, 0x7a, 0xde, 0x1d, 0x63, 0xce, 0xd3, 0xea, 0x2e,
    0x46, 0x67, 0x0d, 0x70, 0xfe, 0x4b, 0xfd, 0xdd, 0x9a, 0x14, 0x28, 0x34,
    0x5c, 0x3c, 0x1c, 0x0c, 0x3a, 0x25, 0xe7, 0x43, 0x95, 0x83, 0xb6, 0x68,
    0x5e, 0x93, 0xe5, 0xa3, 0x2e, 0x95, 0x9c, 0x97, 0x6a, 0x6b, 0xc0, 0xe1,
    0x3a, 0x07, 0xeb, 0x06, 0x1c, 0xa1, 0x92, 0x6f, 0x0d, 0xf5, 0x88, 0x30,
    0x2c, 0x6d, 0x03, 0x36, 0x91, 0x56, 0x3d, 0xaf, 0x7a, 0x34, 0x1c, 0x52,
    0x41, 0xf6, 0x77, 0x5e, 0xdb, 0xb5, 0x28, 0x7e, 0x01, 0xe1, 0x80, 0x53,
    0x89, 0x1e, 0x92, 0x05, 0xcf, 0xe2, 0x64, 0x3b, 0x15, 0xa3, 0x76, 0x91,
    0x38, 0xe9, 0x9f, 0x21, 0xbd, 0x50, 0x8b, 0xc8, 0xa4, 0xe4, 0xd4, 0x1c,
    0xe5, 0xdc, 0x3b, 0xd7, 0xf6, 0x46, 0x2f, 0x47, 0x43, 0x16, 0x97, 0x46,
    0x39, 0x81, 0x64, 0x7d, 0x89, 0xb8, 0xb3, 0xec, 0x90, 0x7e, 0x34, 0x1b,
    0x48, 0xfc, 0x0f, 0xe2, 0xe5, 0x9c, 0x78, 0x64, 0x10, 0x33, 0xc6, 0x91,
    0xef, 0x1d, 0xde, 0xf9, 0x43, 0x7f, 0x45, 0xaf, 0x43, 0x17, 0xc1, 0xe7,
    0xfd, 0xc9, 0x49, 0x7d, 0xb3, 0xfe, 0x85, 0x9e, 0xe2, 0x2e, 0xdc, 0x03,
    0xe3, 0x79, 0x9b, 0x97, 0x4d, 0xb5, 0x6d, 0x21, 0x61, 0x9d, 0x0f, 0x30,
    0x6d, 0x7b, 0x6a, 0xa8, 0xc7, 0x8c, 0x95, 0xdc, 0xe0, 0x5f, 0xb7, 0xe0,
    0xf7, 0x58, 0x9f, 0x3d, 0xef, 0x07, 0xc3, 0xed, 0x9f, 0x4f, 0x0f, 0x5c,
    0x2f, 0x6f, 0x16, 0xe5, 0xa7, 0xb2, 0x1c, 0x4f, 0x7e, 0xee, 0x29, 0xf9,
    0xf6, 0x94, 0xf5, 0x1f, 0x50, 0x4b, 0x03, 0x04, 0x14, 0x00, 0x00, 0x00,
    0x08, 0x00, 0x00, 0x00, 0x21, 0x00, 0xa2, 0xc0, 0x5e, 0x9a, 0x32, 0x01,
    0x00, 0x00, 0x51, 0x02, 0x00, 0x00, 0x11, 0x00, 0x00, 0x00, 0x64, 0x6f,
    0x63, 0x50, 0x72, 0x6f, 0x70, 0x73, 0x2f, 0x63, 0x6f, 0x72, 0x65, 0x2e,
    0x78, 0x6d, 0x6c, 0x7d, 0x92, 0x5f, 0x6b, 0xc3, 0x20, 0x14, 0xc5, 0xdf,
    0x07, 0xfb, 0x0e, 0xc1, 0xf7, 0x44, 0xed, 0x3f, 0x36, 0x49, 0x52, 0xd8,
    0x46, 0x9f, 0x56, 0x18, 0x2c, 0xa3, 0x63, 0x6f, 0xa2, 0xb7, 0xad, 0x2c,
    0x1a, 0x51, 0xb7, 0xb4, 0xdf, 0x7e, 0x26, 0x6d, 0xd3, 0x16, 0xca, 0xc0,
    0x17, 0x3d, 0xe7, 0xfe, 0xee, 0xb9, 0x17, 0xf3, 0xf9, 0x4e, 0xd7, 0xc9,
    0x2f, 0x38, 0xaf, 0x1a, 0x53, 0x20, 0x9a, 0x11, 0x94, 0x80, 0x11, 0x8d,
    0x54, 0x66, 0x53, 0xa0, 0x8f, 0x6a, 0x91, 0x3e, 0xa0, 0xc4, 0x07, 0x6e,
    0x24, 0xaf, 0x1b, 0x03, 0x05, 0xda, 0x83, 0x47, 0xf3, 0xf2, 0xfe, 0x2e,
    0x17, 0x96, 0x89, 0xc6, 0xc1, 0x9b, 0x6b, 0x2c, 0xb8, 0xa0, 0xc0, 0x27,
    0x91, 0x64, 0x3c, 0x13, 0xb6, 0x40, 0xdb, 0x10, 0x2c, 0xc3, 0xd8, 0x8b,
    0x2d, 0x68, 0xee, 0xb3, 0xe8, 0x30, 0x51, 0x5c, 0x37, 0x4e, 0xf3, 0x10,
    0xaf, 0x6e, 0x83, 0x2d, 0x17, 0xdf, 0x7c, 0x03, 0x78, 0x44, 0xc8, 0x0c,
    0x6b, 0x08, 0x5c, 0xf2, 0xc0, 0x71, 0x07, 0x4c, 0xed, 0x40, 0x44, 0x47,
    0xa4, 0x14, 0x03, 0xd2, 0xfe, 0xb8, 0xba, 0x07, 0x48, 0x81, 0xa1, 0x06,
    0x0d, 0x26, 0x78, 0x4c, 0x33, 0x8a, 0xcf, 0xde, 0x00, 0x4e, 0xfb, 0x9b,
    0x05, 0xbd, 0x72, 0xe1, 0xd4, 0x2a, 0xec, 0x2d, 0xdc, 0xb4, 0x9e, 0xc4,
    0xc1, 0xbd, 0xf3, 0x6a, 0x30, 0xb6, 0x6d, 0x9b, 0xb5, 0xe3, 0xde, 0x1a,
    0xf3, 0x53, 0xfc, 0xb9, 0x7c, 0x7d, 0xef, 0x47, 0x4d, 0x95, 0xe9, 0x76,
    0x25, 0x00, 0x95, 0xb9, 0x14, 0x4c, 0x38, 0xe0, 0xa1, 0x71, 0x65, 0x8e,
    0x2f, 0x2f, 0x71, 0x71, 0x35, 0xf7, 0x61, 0x19, 0x77, 0xbc, 0x56, 0x20,
    0x9f, 0xf6, 0x51, 0xbf, 0xf1, 0x76, 0x1c, 0xe4, 0x50, 0x07, 0x32, 0x89,
    0x01, 0xd8, 0x21, 0xee, 0x49, 0x59, 0x8d, 0x9f, 0x5f, 0xaa, 0x05, 0x2a,
    0xbb, 0x1d, 0xa6, 0xe4, 0x31, 0xa5, 0xb3, 0x8a, 0x10, 0xd6, 0x9f, 0xaf,
    0xae, 0xe5, 0x55, 0xfd, 0x19, 0xa8, 0x8f, 0x4d, 0xfe, 0x25, 0xd2, 0x49,
    0x4a, 0xa6, 0x29, 0x1d, 0x55, 0x74, 0xc2, 0xa6, 0x33, 0x46, 0xa7, 0x17,
    0xc4, 0x13, 0xe0, 0x90, 0xfb, 0xfa, 0x13, 0x94, 0x7f, 0x50, 0x4b, 0x03,
    0x04, 0x14, 0x00, 0x00, 0x00, 0x00, 0x00, 0x50, 0x57, 0xac, 0x44, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x09,
    0x00, 0x00, 0x00, 0x78, 0x6c, 0x2f, 0x5f, 0x72, 0x65, 0x6c, 0x73, 0x2f,
    0x50, 0x4b, 0x03, 0x04, 0x14, 0x00, 0x00, 0x00, 0x08, 0x00, 0x4d, 0x57,
    0xac, 0x44, 0xc4, 0x41, 0xfb, 0x70, 0xb7, 0x00, 0x00, 0x00, 0x2c, 0x01,
    0x00, 0x00, 0x1a, 0x00, 0x00, 0x00, 0x78, 0x6c, 0x2f, 0x5f, 0x72, 0x65,
    0x6c, 0x73, 0x2f, 0x77, 0x6f, 0x72, 0x6b, 0x62, 0x6f, 0x6f, 0x6b, 0x2e,
    0x78, 0x6d, 0x6c, 0x2e, 0x72, 0x65, 0x6c, 0x73, 0x8d, 0xcf, 0xcd, 0x0a,
    0xc2, 0x30, 0x0c, 0x07, 0xf0, 0xbb, 0xe0, 0x3b, 0x94, 0xdc, 0x5d, 0x36,
    0x0f, 0x22, 0xb2, 0x6e, 0x17, 0x11, 0x76, 0x95, 0xf9, 0x00, 0xa5, 0xcb,
    0x3e, 0xd8, 0xd6, 0x96, 0xa6, 0x7e, 0xec, 0xed, 0x2d, 0x1e, 0x44, 0xc1,
    0x83, 0xa7, 0x90, 0x84, 0xfc, 0xc2, 0x3f, 0x2f, 0x1f, 0xf3, 0x24, 0x6e,
    0xe4, 0x79, 0xb0, 0x46, 0x42, 0x96, 0xa4, 0x20, 0xc8, 0x68, 0xdb, 0x0c,
    0xa6, 0x93, 0x70, 0xa9, 0x4f, 0x9b, 0x3d, 0x08, 0x0e, 0xca, 0x34, 0x6a,
    0xb2, 0x86, 0x24, 0x2c, 0xc4, 0x50, 0x16, 0xeb, 0x55, 0x7e, 0xa6, 0x49,
    0x85, 0x78, 0xc4, 0xfd, 0xe0, 0x58, 0x44, 0xc5, 0xb0, 0x84, 0x3e, 0x04,
    0x77, 0x40, 0x64, 0xdd, 0xd3, 0xac, 0x38, 0xb1, 0x8e, 0x4c, 0xdc, 0xb4,
    0xd6, 0xcf, 0x2a, 0xc4, 0xd6, 0x77, 0xe8, 0x94, 0x1e, 0x55, 0x47, 0xb8,
    0x4d, 0xd3, 0x1d, 0xfa, 0x4f, 0x03, 0x8a, 0x2f, 0x53, 0x54, 0x8d, 0x04,
    0x5f, 0x35, 0x19, 0x88, 0x7a, 0x71, 0xf4, 0x8f, 0x6d, 0xdb, 0x76, 0xd0,
    0x74, 0xb4, 0xfa, 0x3a, 0x93, 0x09, 0x3f, 0x5e, 0xe0, 0xdd, 0xfa, 0x91,
    0x7b, 0xa2, 0x10, 0x51, 0xe5, 0x3b, 0x0a, 0x12, 0xde, 0x23, 0xc6, 0x57,
    0xc9, 0x92, 0xa8, 0x02, 0x16, 0x39, 0x7e, 0x25, 0x8c, 0x91, 0x9f, 0x50,
    0x4b, 0x03, 0x04, 0x14, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00, 0x21,
    0x00, 0x2c, 0x65, 0xe1, 0xa5, 0x48, 0x01, 0x00, 0x00, 0x26, 0x02, 0x00,
    0x00, 0x0f, 0x00, 0x00, 0x00, 0x78, 0x6c, 0x2f, 0x77, 0x6f, 0x72, 0x6b,
    0x62, 0x6f, 0x6f, 0x6b, 0x2e, 0x78, 0x6d, 0x6c, 0x8d, 0x91, 0xc1, 0x4e,
    0xc3, 0x30, 0x0c, 0x86, 0xef, 0x48, 0xbc, 0x43, 0xe4, 0x3b, 0x6b, 0x9b,
    0x95, 0x69, 0x4c, 0x6b, 0x27, 0x21, 0x40, 0xec, 0x82, 0x76, 0x18, 0xec,
    0x1c, 0x1a, 0x77, 0x8d, 0x96, 0x26, 0x55, 0x92, 0xae, 0xdb, 0xdb, 0xe3,
    0xb6, 0x2a, 0xe3, 0xc8, 0xc9, 0xfe, 0x9d, 0xf8, 0xcb, 0x6f, 0x67, 0xbd,
    0xb9, 0xd4, 0x9a, 0x9d, 0xd1, 0x79, 0x65, 0x4d, 0x06, 0xc9, 0x2c, 0x06,
    0x86, 0xa6, 0xb0, 0x52, 0x99, 0x63, 0x06, 0x9f, 0xfb, 0xb7, 0x87, 0x25,
    0x30, 0x1f, 0x84, 0x91, 0x42, 0x5b, 0x83, 0x19, 0x5c, 0xd1, 0xc3, 0x26,
    0xbf, 0xbf, 0x5b, 0x77, 0xd6, 0x9d, 0xbe, 0xad, 0x3d, 0x31, 0x02, 0x18,
    0x9f, 0x41, 0x15, 0x42, 0xb3, 0x8a, 0x22, 0x5f, 0x54, 0x58, 0x0b, 0x3f,
    0xb3, 0x0d, 0x1a, 0x3a, 0x29, 0xad, 0xab, 0x45, 0x20, 0xe9, 0x8e, 0x91,
    0x6f, 0x1c, 0x0a, 0xe9, 0x2b, 0xc4, 0x50, 0xeb, 0x88, 0xc7, 0xf1, 0x22,
    0xaa, 0x85, 0x32, 0x30, 0x12, 0x56, 0xee, 0x3f, 0x0c, 0x5b, 0x96, 0xaa,
    0xc0, 0x17, 0x5b, 0xb4, 0x35, 0x9a, 0x30, 0x42, 0x1c, 0x6a, 0x11, 0xc8,
    0xbe, 0xaf, 0x54, 0xe3, 0x21, 0x5f, 0x97, 0x4a, 0xe3, 0xd7, 0x38, 0x11,
    0x13, 0x4d, 0xf3, 0x21, 0x6a, 0xf2, 0x7d, 0xd1, 0xc0, 0xb4, 0xf0, 0xe1,
    0x55, 0xaa, 0x80, 0x32, 0x83, 0x47, 0x92, 0xb6, 0xc3, 0x5b, 0x21, 0x05,
    0xe6, 0xda, 0xe6, 0xb9, 0x55, 0x9a, 0xc4, 0xd3, 0x3c, 0x9e, 0x43, 0x94,
    0xff, 0x0e, 0xb9, 0x73, 0x8c, 0xa8, 0x01, 0xdd, 0xce, 0xa9, 0xb3, 0x28,
    0xae, 0xb4, 0x29, 0x60, 0x12, 0x4b, 0xd1, 0xea, 0xb0, 0x27, 0xb3, 0xd3,
    0x7b, 0x54, 0xe7, 0x29, 0xe7, 0x8b, 0xbe, 0xb7, 0xef, 0xfb, 0x52, 0xd8,
    0xf9, 0x1b, 0xa6, 0x97, 0xec, 0x72, 0x50, 0x46, 0xda, 0x2e, 0x03, 0x9e,
    0xd2, 0xb2, 0xaf, 0x93, 0x4a, 0x62, 0xb2, 0xd4, 0x0d, 0xe2, 0xa0, 0x64,
    0xa8, 0xa8, 0x92, 0x2e, 0x6f, 0xb5, 0x77, 0x54, 0xc7, 0x2a, 0x64, 0xb0,
    0x8c, 0x93, 0xb8, 0xa7, 0x47, 0x7f, 0xf0, 0xc3, 0x4a, 0xa7, 0xc8, 0xcc,
    0x30, 0xef, 0x90, 0xd3, 0xd7, 0xf5, 0x61, 0x2b, 0x07, 0xbf, 0x6e, 0xa5,
    0x28, 0x71, 0x5b, 0x99, 0x0c, 0x80, 0xa9, 0xab, 0x10, 0xba, 0xa0, 0xf9,
    0xfa, 0x30, 0x5c, 0xe4, 0x9c, 0x27, 0xe3, 0x8d, 0xc9, 0x76, 0xfe, 0x03,
    0x50, 0x4b, 0x03, 0x04, 0x14, 0x00, 0x00, 0x00, 0x00, 0x00, 0x11, 0x57,
    0xac, 0x44, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x0e, 0x00, 0x00, 0x00, 0x78, 0x6c, 0x2f, 0x77, 0x6f, 0x72,
    0x6b, 0x73, 0x68, 0x65, 0x65, 0x74, 0x73, 0x2f, 0x50, 0x4b, 0x03, 0x04,
    0x14, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00, 0x21, 0x00, 0xe6, 0x55,
    0xa8, 0xe3, 0x5d, 0x01, 0x00, 0x00, 0x84, 0x02, 0x00, 0x00, 0x18, 0x00,
    0x00, 0x00, 0x78, 0x6c, 0x2f, 0x77, 0x6f, 0x72, 0x6b, 0x73, 0x68, 0x65,
    0x65, 0x74, 0x73, 0x2f, 0x73, 0x68, 0x65, 0x65, 0x74, 0x31, 0x2e, 0x78,
    0x6d, 0x6c, 0x8d, 0x92, 0x4f, 0x6b, 0x02, 0x31, 0x10, 0xc5, 0xef, 0x85,
    0x7e, 0x87, 0x90, 0xbb, 0x46, 0x6d, 0x6d, 0xab, 0xb8, 0x4a, 0x41, 0xa4,
    0x1e, 0x0a, 0xa5, 0xff, 0xee, 0xd9, 0xec, 0xec, 0x6e, 0x30, 0xc9, 0x2c,
    0xc9, 0x58, 0xf5, 0xdb, 0x77, 0x76, 0xad, 0x52, 0xf0, 0xe2, 0x6d, 0x5e,
    0x26, 0xf3, 0xe3, 0xbd, 0x49, 0x66, 0x8b, 0xbd, 0x77, 0xe2, 0x07, 0x62,
    0xb2, 0x18, 0x32, 0x39, 0xec, 0x0f, 0xa4, 0x80, 0x60, 0xb0, 0xb0, 0xa1,
    0xca, 0xe4, 0xd7, 0xe7, 0xaa, 0xf7, 0x24, 0x45, 0x22, 0x1d, 0x0a, 0xed,
    0x30, 0x40, 0x26, 0x0f, 0x90, 0xe4, 0x62, 0x7e, 0x7b, 0x33, 0xdb, 0x61,
    0xdc, 0xa4, 0x1a, 0x80, 0x04, 0x13, 0x42, 0xca, 0x64, 0x4d, 0xd4, 0x4c,
    0x95, 0x4a, 0xa6, 0x06, 0xaf, 0x53, 0x1f, 0x1b, 0x08, 0xdc, 0x29, 0x31,
    0x7a, 0x4d, 0x2c, 0x63, 0xa5, 0x52, 0x13, 0x41, 0x17, 0xdd, 0x90, 0x77,
    0x6a, 0x34, 0x18, 0x3c, 0x28, 0xaf, 0x6d, 0x90, 0x47, 0xc2, 0x34, 0x5e,
    0xc3, 0xc0, 0xb2, 0xb4, 0x06, 0x96, 0x68, 0xb6, 0x1e, 0x02, 0x1d, 0x21,
    0x11, 0x9c, 0x26, 0xf6, 0x9f, 0x6a, 0xdb, 0xa4, 0x13, 0xcd, 0x9b, 0x6b,
    0x70, 0x5e, 0xc7, 0xcd, 0xb6, 0xe9, 0x19, 0xf4, 0x0d, 0x23, 0x72, 0xeb,
    0x2c, 0x1d, 0x3a, 0xa8, 0x14, 0xde, 0x4c, 0xd7, 0x55, 0xc0, 0xa8, 0x73,
    0xc7, 0xb9, 0xf7, 0xc3, 0x7b, 0x6d, 0x4e, 0xec, 0x4e, 0x5c, 0xe0, 0xbd,
    0x35, 0x11, 0x13, 0x96, 0xd4, 0x67, 0xdc, 0x9f, 0xd1, 0xcb, 0xcc, 0x13,
    0x35, 0x51, 0x4c, 0x9a, 0xcf, 0x0a, 0xcb, 0x09, 0xda, 0xb5, 0x8b, 0x08,
    0x65, 0x26, 0x9f, 0x87, 0x52, 0xcd, 0x67, 0xdd, 0xc5, 0x6f, 0x0b, 0xbb,
    0xf4, 0xaf, 0x16, 0xa4, 0xf3, 0x0f, 0x70, 0x60, 0x08, 0x0a, 0x7e, 0x23,
    0x29, 0xda, 0xdd, 0xe7, 0x88, 0x9b, 0xb6, 0xb9, 0xe6, 0xa3, 0x41, 0x3b,
    0xaa, 0x2e, 0x66, 0x57, 0x5d, 0xd0, 0xb7, 0x28, 0x0a, 0x28, 0xf5, 0xd6,
    0xd1, 0x3b, 0xee, 0x5e, 0xc0, 0x56, 0x35, 0x31, 0x64, 0xcc, 0x59, 0xda,
    0x14, 0xd3, 0xe2, 0xb0, 0x84, 0x64, 0x78, 0x97, 0x8c, 0xe9, 0x8f, 0xc6,
    0x67, 0x13, 0x4b, 0x4d, 0x9a, 0xeb, 0x46, 0x57, 0xf0, 0xaa, 0x63, 0x65,
    0x43, 0x12, 0x0e, 0xca, 0xee, 0xd6, 0xa3, 0x14, 0xf1, 0x88, 0xe9, 0x6a,
    0xc2, 0xa6, 0xab, 0x18, 0x99, 0x23, 0x11, 0xfa, 0x93, 0xaa, 0x39, 0x39,
    0xc4, 0x56, 0xdd, 0x49, 0x51, 0x22, 0xd2, 0x49, 0xb4, 0x6e, 0xcf, 0xff,
    0x67, 0xfe, 0x0b, 0x50, 0x4b, 0x03, 0x04, 0x14, 0x00, 0x00, 0x00, 0x08,
    0x00, 0x65, 0x57, 0xac, 0x44, 0xba, 0x83, 0x84, 0x3d, 0x2a, 0x01, 0x00,
    0x00, 0x1f, 0x03, 0x00, 0x00, 0x13, 0x00, 0x00, 0x00, 0x5b, 0x43, 0x6f,
    0x6e, 0x74, 0x65, 0x6e, 0x74, 0x5f, 0x54, 0x79, 0x70, 0x65, 0x73, 0x5d,
    0x2e, 0x78, 0x6d, 0x6c, 0xad, 0x92, 0xcd, 0x6e, 0xc2, 0x30, 0x10, 0x84,
    0xef, 0x95, 0xfa, 0x0e, 0x96, 0xaf, 0x28, 0x36, 0xf4, 0x50, 0x55, 0x15,
    0x81, 0x43, 0x7f, 0x8e, 0x2d, 0x07, 0xfa, 0x00, 0xae, 0xbd, 0x49, 0x2c,
    0xfc, 0x27, 0xaf, 0xa1, 0xf0, 0xf6, 0xdd, 0x04, 0xda, 0x03, 0xa2, 0xb4,
    0x48, 0x3d, 0x59, 0xc9, 0xce, 0xcc, 0x37, 0x89, 0x77, 0x3a, 0xdf, 0x7a,
    0xc7, 0x36, 0x90, 0xd1, 0xc6, 0x50, 0xf3, 0x89, 0x18, 0x73, 0x06, 0x41,
    0x47, 0x63, 0x43, 0x5b, 0xf3, 0xb7, 0xe5, 0x73, 0x75, 0xc7, 0x19, 0x16,
    0x15, 0x8c, 0x72, 0x31, 0x40, 0xcd, 0x77, 0x80, 0x7c, 0x3e, 0xbb, 0xbe,
    0x9a, 0x2e, 0x77, 0x09, 0x90, 0x91, 0x3b, 0x60, 0xcd, 0xbb, 0x52, 0xd2,
    0xbd, 0x94, 0xa8, 0x3b, 0xf0, 0x0a, 0x45, 0x4c, 0x10, 0x68, 0xd2, 0xc4,
    0xec, 0x55, 0xa1, 0xc7, 0xdc, 0xca, 0xa4, 0xf4, 0x4a, 0xb5, 0x20, 0x6f,
    0xc6, 0xe3, 0x5b, 0xa9, 0x63, 0x28, 0x10, 0x4a, 0x55, 0xfa, 0x0c, 0x3e,
    0x9b, 0x3e, 0x42, 0xa3, 0xd6, 0xae, 0xb0, 0xa7, 0x2d, 0xbd, 0xde, 0x37,
    0xc9, 0xe0, 0x90, 0xb3, 0x87, 0xbd, 0xb0, 0x67, 0xd5, 0x5c, 0xa5, 0xe4,
    0xac, 0x56, 0x85, 0xe6, 0x72, 0x13, 0xcc, 0x11, 0xa5, 0x3a, 0x10, 0x04,
    0x39, 0x07, 0x0d, 0x76, 0x36, 0xe1, 0x88, 0x04, 0x5c, 0x9e, 0x24, 0xf4,
    0x93, 0x9f, 0x01, 0x07, 0xdf, 0x2b, 0xfd, 0x9a, 0x6c, 0x0d, 0xb0, 0x85,
    0xca, 0xe5, 0x45, 0x79, 0x52, 0xc9, 0xad, 0x93, 0x1f, 0x31, 0xaf, 0xde,
    0x63, 0x5c, 0x89, 0xf3, 0x21, 0x27, 0x5a, 0xc6, 0xa6, 0xb1, 0x1a, 0x4c,
    0xd4, 0x6b, 0x4f, 0x16, 0x81, 0x29, 0x83, 0x32, 0xd8, 0x01, 0x14, 0xef,
    0xc4, 0x70, 0x0a, 0xaf, 0x6c, 0x18, 0xfd, 0xce, 0x1f, 0xc4, 0x28, 0x87,
    0x63, 0xf2, 0xcf, 0x45, 0xbe, 0xf3, 0xcf, 0xf5, 0x20, 0xef, 0x22, 0xc7,
    0x84, 0x74, 0x9d, 0x19, 0x2e, 0x2f, 0xf0, 0x75, 0x5f, 0xbd, 0xbb, 0x4a,
    0x14, 0x04, 0xb9, 0x58, 0xc0, 0x3f, 0x11, 0x29, 0xfa, 0x72, 0xe0, 0xd1,
    0x17, 0x43, 0xbf, 0x0a, 0x06, 0xcc, 0x09, 0xb6, 0x1c, 0x96, 0x9b, 0xb6,
    0xfc, 0x13, 0x50, 0x4b, 0x03, 0x04, 0x14, 0x00, 0x00, 0x00, 0x08, 0x00,
    0x00, 0x00, 0x21, 0x00, 0xb5, 0x55, 0x30, 0x23, 0xec, 0x00, 0x00, 0x00,
    0x4c, 0x02, 0x00, 0x00, 0x0b, 0x00, 0x00, 0x00, 0x5f, 0x72, 0x65, 0x6c,
    0x73, 0x2f, 0x2e, 0x72, 0x65, 0x6c, 0x73, 0x8d, 0x92, 0xcd, 0x4e, 0xc3,
    0x30, 0x0c, 0x80, 0xef, 0x48, 0xbc, 0x43, 0xe4, 0xfb, 0xea, 0x6e, 0x48,
    0x08, 0xa1, 0xa5, 0xbb, 0x20, 0xa4, 0xdd, 0x10, 0x2a, 0x0f, 0x60, 0x12,
    0xf7, 0x47, 0x6d, 0xe3, 0x28, 0x09, 0xd0, 0xbd, 0x3d, 0xe1, 0x80, 0xa0,
    0xd2, 0x18, 0x3d, 0xc6, 0xb1, 0x3f, 0x7f, 0xb6, 0xbc, 0x3f, 0xcc, 0xd3,
    0xa8, 0xde, 0x39, 0xc4, 0x5e, 0x9c, 0x86, 0x6d, 0x51, 0x82, 0x62, 0x67,
    0xc4, 0xf6, 0xae, 0xd5, 0xf0, 0x52, 0x3f, 0x6e, 0xee, 0x40, 0xc5, 0x44,
    0xce, 0xd2, 0x28, 0x8e, 0x35, 0x9c, 0x38, 0xc2, 0xa1, 0xba, 0xbe, 0xda,
    0x3f, 0xf3, 0x48, 0x29, 0x17, 0xc5, 0xae, 0xf7, 0x51, 0x65, 0x8a, 0x8b,
    0x1a, 0xba, 0x94, 0xfc, 0x3d, 0x62, 0x34, 0x1d, 0x4f, 0x14, 0x0b, 0xf1,
    0xec, 0xf2, 0x4f, 0x23, 0x61, 0xa2, 0x94, 0x9f, 0xa1, 0x45, 0x4f, 0x66,
    0xa0, 0x96, 0x71, 0x57, 0x96, 0xb7, 0x18, 0x7e, 0x33, 0xa0, 0x5a, 0x30,
    0xd5, 0xd1, 0x6a, 0x08, 0x47, 0x7b, 0x03, 0xaa, 0x3e, 0x79, 0x5e, 0xc3,
    0x96, 0xa6, 0xe9, 0x0d, 0x3f, 0x88, 0x79, 0x9b, 0xd8, 0xa5, 0x33, 0x2d,
    0x90, 0xe7, 0xc4, 0xce, 0xb2, 0xdd, 0xf8, 0x90, 0xeb, 0x43, 0xea, 0xf3,
    0x34, 0xaa, 0xa6, 0xd0, 0x72, 0xd2, 0x60, 0xc5, 0x3c, 0xe5, 0x70, 0x44,
    0xf2, 0xbe, 0xc8, 0x68, 0xc0, 0xf3, 0x46, 0xbb, 0xf5, 0x46, 0x7f, 0x4f,
    0x8b, 0x13, 0x27, 0xb2, 0x94, 0x08, 0x8d, 0x04, 0xbe, 0xec, 0xf3, 0x95,
    0x71, 0x49, 0x68, 0xbb, 0x5e, 0xe8, 0xff, 0x15, 0x2d, 0x33, 0x7e, 0x6c,
    0xe6, 0x11, 0x3f, 0x24, 0x0c, 0xaf, 0x22, 0xc3, 0xb7, 0x0b, 0x2e, 0x6e,
    0xa0, 0xfa, 0x04, 0x50, 0x4b, 0x01, 0x02, 0x14, 0x00, 0x14, 0x00, 0x00,
    0x00, 0x08, 0x00, 0x00, 0x00, 0x21, 0x00, 0xf8, 0x17, 0x86, 0x86, 0x7a,
    0x01, 0x00, 0x00, 0x10, 0x03, 0x00, 0x00, 0x10, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x20, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x64, 0x6f, 0x63, 0x50, 0x72, 0x6f, 0x70, 0x73, 0x2f, 0x61, 0x70,
    0x70, 0x2e, 0x78, 0x6d, 0x6c, 0x50, 0x4b, 0x01, 0x02, 0x14, 0x00, 0x14,
    0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00, 0x21, 0x00, 0xa2, 0xc0, 0x5e,
    0x9a, 0x32, 0x01, 0x00, 0x00, 0x51, 0x02, 0x00, 0x00, 0x11, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x20, 0x00, 0x00, 0x00, 0xa8,
    0x01, 0x00, 0x00, 0x64, 0x6f, 0x63, 0x50, 0x72, 0x6f, 0x70, 0x73, 0x2f,
    0x63, 0x6f, 0x72, 0x65, 0x2e, 0x78, 0x6d, 0x6c, 0x50, 0x4b, 0x01, 0x02,
    0x14, 0x00, 0x14, 0x00, 0x00, 0x00, 0x00, 0x00, 0x50, 0x57, 0xac, 0x44,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x09, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x00,
    0x00, 0x00, 0x09, 0x03, 0x00, 0x00, 0x78, 0x6c, 0x2f, 0x5f, 0x72, 0x65,
    0x6c, 0x73, 0x2f, 0x50, 0x4b, 0x01, 0x02, 0x14, 0x00, 0x14, 0x00, 0x00,
    0x00, 0x08, 0x00, 0x4d, 0x57, 0xac, 0x44, 0xc4, 0x41, 0xfb, 0x70, 0xb7,
    0x00, 0x00, 0x00, 0x2c, 0x01, 0x00, 0x00, 0x1a, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x20, 0x00, 0x00, 0x00, 0x30, 0x03, 0x00,
    0x00, 0x78, 0x6c, 0x2f, 0x5f, 0x72, 0x65, 0x6c, 0x73, 0x2f, 0x77, 0x6f,
    0x72, 0x6b, 0x62, 0x6f, 0x6f, 0x6b, 0x2e, 0x78, 0x6d, 0x6c, 0x2e, 0x72,
    0x65, 0x6c, 0x73, 0x50, 0x4b, 0x01, 0x02, 0x14, 0x00, 0x14, 0x00, 0x00,
    0x00, 0x08, 0x00, 0x00, 0x00, 0x21, 0x00, 0x2c, 0x65, 0xe1, 0xa5, 0x48,
    0x01, 0x00, 0x00, 0x26, 0x02, 0x00, 0x00, 0x0f, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x20, 0x00, 0x00, 0x00, 0x1f, 0x04, 0x00,
    0x00, 0x78, 0x6c, 0x2f, 0x77, 0x6f, 0x72, 0x6b, 0x62, 0x6f, 0x6f, 0x6b,
    0x2e, 0x78, 0x6d, 0x6c, 0x50, 0x4b, 0x01, 0x02, 0x14, 0x00, 0x14, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x11, 0x57, 0xac, 0x44, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x0e, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x00, 0x00, 0x00, 0x94, 0x05,
    0x00, 0x00, 0x78, 0x6c, 0x2f, 0x77, 0x6f, 0x72, 0x6b, 0x73, 0x68, 0x65,
    0x65, 0x74, 0x73, 0x2f, 0x50, 0x4b, 0x01, 0x02, 0x14, 0x00, 0x14, 0x00,
    0x00, 0x00, 0x08, 0x00, 0x00, 0x00, 0x21, 0x00, 0xe6, 0x55, 0xa8, 0xe3,
    0x5d, 0x01, 0x00, 0x00, 0x84, 0x02, 0x00, 0x00, 0x18, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x20, 0x00, 0x00, 0x00, 0xc0, 0x05,
    0x00, 0x00, 0x78, 0x6c, 0x2f, 0x77, 0x6f, 0x72, 0x6b, 0x73, 0x68, 0x65,
    0x65, 0x74, 0x73, 0x2f, 0x73, 0x68, 0x65, 0x65, 0x74, 0x31, 0x2e, 0x78,
    0x6d, 0x6c, 0x50, 0x4b, 0x01, 0x02, 0x14, 0x00, 0x14, 0x00, 0x00, 0x00,
    0x08, 0x00, 0x65, 0x57, 0xac, 0x44, 0xba, 0x83, 0x84, 0x3d, 0x2a, 0x01,
    0x00, 0x00, 0x1f, 0x03, 0x00, 0x00, 0x13, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x01, 0x00, 0x20, 0x00, 0x00, 0x00, 0x53, 0x07, 0x00, 0x00,
    0x5b, 0x43, 0x6f, 0x6e, 0x74, 0x65, 0x6e, 0x74, 0x5f, 0x54, 0x79, 0x70,
    0x65, 0x73, 0x5d, 0x2e, 0x78, 0x6d, 0x6c, 0x50, 0x4b, 0x01, 0x02, 0x14,
    0x00, 0x14, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00, 0x21, 0x00, 0xb5,
    0x55, 0x30, 0x23, 0xec, 0x00, 0x00, 0x00, 0x4c, 0x02, 0x00, 0x00, 0x0b,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x20, 0x00, 0x00,
    0x00, 0xae, 0x08, 0x00, 0x00, 0x5f, 0x72, 0x65, 0x6c, 0x73, 0x2f, 0x2e,
    0x72, 0x65, 0x6c, 0x73, 0x50, 0x4b, 0x05, 0x06, 0x00, 0x00, 0x00, 0x00,
    0x09, 0x00, 0x09, 0x00, 0x35, 0x02, 0x00, 0x00, 0xc3, 0x09, 0x00, 0x00,
    0x00, 0x00
};

} // namespace

#ifdef _WIN32
#include <Windows.h>
#include <Shlwapi.h>
void file::copy(const std::string &source, const std::string &destination, bool overwrite)
{
    assert(source.size() + 1 < MAX_PATH);
    assert(destination.size() + 1 < MAX_PATH);

    std::wstring source_wide(source.begin(), source.end());
    std::wstring destination_wide(destination.begin(), destination.end());

    BOOL result = CopyFile(source_wide.c_str(), destination_wide.c_str(), !overwrite);

    if(result == 0)
    {
        switch(GetLastError())
        {
        case ERROR_ACCESS_DENIED: throw std::runtime_error("Access is denied");
        case ERROR_ENCRYPTION_FAILED: throw std::runtime_error("The specified file could not be encrypted");
        case ERROR_FILE_NOT_FOUND: throw std::runtime_error("The source file wasn't found");
        default:
            if(!overwrite)
            {
                throw std::runtime_error("The destination file already exists");
            }
            throw std::runtime_error("Unknown error");
        }
    }
}

bool file::exists(const std::string &path)
{
    std::wstring path_wide(path.begin(), path.end());
    return PathFileExists(path_wide.c_str()) && !PathIsDirectory(path_wide.c_str());
}

#else

#include <sys/stat.h>
void file::copy(const std::string &source, const std::string &destination, bool overwrite)
{
    if(!overwrite && exists(destination))
    {
	throw std::runtime_error("destination file already exists and overwrite==false");
    }

    std::ifstream src(source, std::ios::binary);
    std::ofstream dst(destination, std::ios::binary);
    
    dst << src.rdbuf();
}

bool file::exists(const std::string &path)
{
    struct stat fileAtt;
    
    if (stat(path.c_str(), &fileAtt) != 0)
    {
	throw std::runtime_error("stat failed");
    }
    
    return S_ISREG(fileAtt.st_mode);
}

#endif //_WIN32


zip_file::zip_file(const std::string &filename, file_mode mode, file_access access)
: filename_(filename),
mode_(mode),
access_(access),
zip_file_(nullptr),
unzip_file_(nullptr),
modified_(false),
current_state_(state::closed)
{
    switch(mode)
    {
    case file_mode::open:
        read_all();
        break;
    case file_mode::open_or_create:
        if(file_exists(filename))
        {
            read_all();
        }
        else
        {
            flush(true);
        }
        break;
    case file_mode::create:
        flush(true);
        break;
    case file_mode::create_new:
        if(file_exists(filename))
        {
            throw std::runtime_error("file exists");
        }
        flush(true);
        break;
    case file_mode::truncate:
        if((int)access & (int)file_access::read)
        {
            throw std::runtime_error("cannot read from file opened with file_mode truncate");
        }
        flush(true);
        break;
    case file_mode::append:
        read_all();
        break;
    }
}

zip_file::~zip_file()
{
    change_state(state::closed);
}

std::string zip_file::get_file_contents(const std::string &filename)
{
    return files_[filename];
}

void zip_file::set_file_contents(const std::string &filename, const std::string &contents)
{
    if(!has_file(filename) || files_[filename] != contents)
    {
        modified_ = true;
    }

    files_[filename] = contents;
}

void zip_file::delete_file(const std::string &filename)
{
    files_.erase(filename);
}

bool zip_file::has_file(const std::string &filename)
{
    return files_.find(filename) != files_.end();
}

void zip_file::flush(bool force_write)
{
    if(modified_ || force_write)
    {
        write_all();
    }
}

void zip_file::read_all()
{
    if(!((int)access_ & (int)file_access::read))
    {
        throw std::runtime_error("don't have read access");
    }

    change_state(state::read);

    int result = unzGoToFirstFile(unzip_file_);

    std::array<char, 1000> file_name_buffer = {'\0'};
    std::vector<char> file_buffer;

    while(result == UNZ_OK)
    {
        unz_file_info file_info;
        file_name_buffer.fill('\0');

        result = unzGetCurrentFileInfo(unzip_file_, &file_info, file_name_buffer.data(), 
            static_cast<uLong>(file_name_buffer.size()), nullptr, 0, nullptr, 0);

        if(result != UNZ_OK)
        {
            throw result;
        }

        result = unzOpenCurrentFile(unzip_file_);

        if(result != UNZ_OK)
        {
            throw result;
        }

        if(file_buffer.size() < file_info.uncompressed_size + 1)
        {
            file_buffer.resize(file_info.uncompressed_size + 1);
        }
        file_buffer[file_info.uncompressed_size] = '\0';

        result = unzReadCurrentFile(unzip_file_, file_buffer.data(), file_info.uncompressed_size);

        if(result != static_cast<int>(file_info.uncompressed_size))
        {
            throw result;
        }

        std::string current_filename(file_name_buffer.begin(), file_name_buffer.begin() + file_info.size_filename);
        std::string contents(file_buffer.begin(), file_buffer.begin() + file_info.uncompressed_size);

        if(current_filename.back() != '/')
        {
            files_[current_filename] = contents;
        }
        else
        {
            directories_.push_back(current_filename);
        }

        result = unzCloseCurrentFile(unzip_file_);

        if(result != UNZ_OK)
        {
            throw result;
        }

        result = unzGoToNextFile(unzip_file_);
    }
}

void zip_file::write_all()
{
    if(!((int)access_ & (int)file_access::write))
    {
        throw std::runtime_error("don't have write access");
    }

    change_state(state::write, false);

    for(auto file : files_)
    {
        write_to_zip(file.first, file.second, true);
    }

    modified_ = false;
}

std::string zip_file::read_from_zip(const std::string &filename)
{
    if(!((int)access_ & (int)file_access::read))
    {
        throw std::runtime_error("don't have read access");
    }

    change_state(state::read);

    auto result = unzLocateFile(unzip_file_, filename.c_str(), 1);

    if(result != UNZ_OK)
    {
        throw result;
    }

    result = unzOpenCurrentFile(unzip_file_);

    if(result != UNZ_OK)
    {
        throw result;
    }

    unz_file_info file_info;
    std::array<char, 1000> file_name_buffer;
    std::array<char, 1000> extra_field_buffer;
    std::array<char, 1000> comment_buffer;

    unzGetCurrentFileInfo(unzip_file_, &file_info,
        file_name_buffer.data(), static_cast<uLong>(file_name_buffer.size()),
        extra_field_buffer.data(), static_cast<uLong>(extra_field_buffer.size()),
        comment_buffer.data(), static_cast<uLong>(comment_buffer.size()));

    std::vector<char> file_buffer(file_info.uncompressed_size + 1, '\0');
    result = unzReadCurrentFile(unzip_file_, file_buffer.data(), file_info.uncompressed_size);

    if(result != static_cast<int>(file_info.uncompressed_size))
    {
        throw result;
    }

    result = unzCloseCurrentFile(unzip_file_);

    if(result != UNZ_OK)
    {
        throw result;
    }

    return std::string(file_buffer.begin(), file_buffer.end());
}

void zip_file::write_to_zip(const std::string &filename, const std::string &content, bool append)
{
    if(!((int)access_ & (int)file_access::write))
    {
        throw std::runtime_error("don't have write access");
    }

    change_state(state::write, append);

    zip_fileinfo file_info = {0};

    int result = zipOpenNewFileInZip(zip_file_, filename.c_str(), &file_info, nullptr, 0, nullptr, 0, nullptr, Z_DEFLATED, Z_DEFAULT_COMPRESSION);

    if(result != UNZ_OK)
    {
        throw result;
    }

    result = zipWriteInFileInZip(zip_file_, content.data(), static_cast<int>(content.size()));

    if(result != UNZ_OK)
    {
        throw result;
    }

    result = zipCloseFileInZip(zip_file_);

    if(result != UNZ_OK)
    {
        throw result;
    }
}

void zip_file::change_state(state new_state, bool append)
{
    if(new_state == current_state_ && append)
    {
        return;
    }

    switch(new_state)
    {
    case state::closed:
        if(current_state_ == state::write)
        {
            stop_write();
        }
        else if(current_state_ == state::read)
        {
            stop_read();
        }
        break;
    case state::read:
        if(current_state_ == state::write)
        {
            stop_write();
        }
        start_read();
        break;
    case state::write:
        if(current_state_ == state::read)
        {
            stop_read();
        }
        if(current_state_ != state::write)
        {
            start_write(append);
        }
        break;
    default:
        throw std::runtime_error("bad enum");
    }

    current_state_ = new_state;
}

bool zip_file::file_exists(const std::string& name)
{
    std::ifstream f(name.c_str());
    return f.good();
}

void zip_file::start_read()
{
    if(unzip_file_ != nullptr || zip_file_ != nullptr)
    {
        throw std::runtime_error("bad state");
    }

    unzip_file_ = unzOpen(filename_.c_str());

    if(unzip_file_ == nullptr)
    {
        throw std::runtime_error("bad or non-existant file");
    }
}

void zip_file::stop_read()
{
    if(unzip_file_ == nullptr)
    {
        throw std::runtime_error("bad state");
    }

    int result = unzClose(unzip_file_);

    if(result != UNZ_OK)
    {
        throw result;
    }

    unzip_file_ = nullptr;
}

void zip_file::start_write(bool append)
{
    if(unzip_file_ != nullptr || zip_file_ != nullptr)
    {
        throw std::runtime_error("bad state");
    }

    int append_status;

    if(append)
    {
        if(!file_exists(filename_))
        {
            throw std::runtime_error("can't append to non-existent file");
        }
        append_status = APPEND_STATUS_ADDINZIP;
    }
    else
    {
        append_status = APPEND_STATUS_CREATE;
    }

    zip_file_ = zipOpen(filename_.c_str(), append_status);

    if(zip_file_ == nullptr)
    {
        if(append)
        {
            throw std::runtime_error("couldn't append to zip file");
        }
        else
        {
            throw std::runtime_error("couldn't create zip file");
        }
    }
}

void zip_file::stop_write()
{
    if(zip_file_ == nullptr)
    {
        throw std::runtime_error("bad state");
    }

    flush();

    int result = zipClose(zip_file_, nullptr);

    if(result != UNZ_OK)
    {
        throw result;
    }

    zip_file_ = nullptr;
}

const xlnt::color xlnt::color::black(0);
const xlnt::color xlnt::color::white(1);

struct cell_struct
{
    cell_struct(worksheet_struct *ws, const std::string &column, int row)
        : type(cell::type::null), parent_worksheet(ws),
        column(xlnt::cell::column_index_from_string(column) - 1), row(row),
        hyperlink_rel("hyperlink")
    {

    }

    std::string to_string() const;

    cell::type type;

    union
    {
        long double numeric_value;
        bool bool_value;
    };

    
    std::string error_value; 
    tm date_value;
    std::string string_value;
    std::string formula_value;
    worksheet_struct *parent_worksheet;
    int column;
    int row;
    style style;
    relationship hyperlink_rel;
    bool merged;
};

const std::unordered_map<std::string, int> cell::ErrorCodes =
{
    {"#NULL!", 0},
    {"#DIV/0!", 1},
    {"#VALUE!", 2},
    {"#REF!", 3},
    {"#NAME?", 4},
    {"#NUM!", 5},
    {"#N/A!", 6}
};

cell::cell() : root_(nullptr)
{
}

cell::cell(worksheet &worksheet, const std::string &column, int row) : root_(nullptr)
{
    cell self = worksheet.cell(column + std::to_string(row));
    root_ = self.root_;
}


cell::cell(worksheet &worksheet, const std::string &column, int row, const std::string &initial_value) : root_(nullptr)
{
    cell self = worksheet.cell(column + std::to_string(row));
    root_ = self.root_;
    *this = initial_value;
}

cell::cell(cell_struct *root) : root_(root)
{
}

int cell::get_row() const
{
    return root_->row;
}

std::string cell::get_column() const
{
    return get_column_letter(root_->column);
}

cell::type cell::data_type_for_value(const std::string &value)
{
    if(value.empty())
    {
        return type::null;
    }

    if(value[0] == '=')
    {
        return type::formula;
    }
    else if(value[0] == '0')
    {
        if(value.length() > 1)
        {
            if(value[1] == '.' || (value.length() > 2 && (value[1] == 'e' || value[1] == 'E')))
            {
                auto first_non_number = std::find_if(value.begin() + 2, value.end(),
                    [](char c) { return !std::isdigit(c, std::locale::classic()); });
                if(first_non_number == value.end())
                {
                    return type::numeric;
                }
            }
            return type::string;
        }
        return type::numeric;
    }
    else if(value[0] == '#')
    {
        return type::error;
    }
    else
    {
        char *p;
        strtod(value.c_str(), &p);
        if(*p != 0)
        {
            static const std::vector<std::string> possible_booleans = {"TRUE", "true", "FALSE", "false"};
            if(std::find(possible_booleans.begin(), possible_booleans.end(), value) != possible_booleans.end())
            {
                return type::boolean;
            }
            return type::string;
        }
        else
        {
            return type::numeric;
        }
    }
}

void cell::set_explicit_value(const std::string &value, type data_type)
{
    root_->type = data_type;
    switch(data_type)
    {
    case type::formula: root_->formula_value = value; return;
    case type::date: root_->date_value.tm_hour = std::stoi(value); return;
    case type::error: root_->error_value = value; return;
    case type::boolean: root_->bool_value = value == "true"; return;
    case type::null: return;
    case type::numeric: root_->numeric_value = std::stod(value); return;
    case type::string: root_->string_value = value; return;
    }
}

void cell::set_hyperlink(const std::string &hyperlink)
{
    root_->hyperlink_rel = worksheet(root_->parent_worksheet).create_relationship(hyperlink);
}

void cell::set_merged(bool merged)
{
    root_->merged = merged;
}

bool cell::bind_value()
{
    root_->type = type::null;
    return true;
}

bool cell::bind_value(int value)
{
    root_->type = type::numeric;
    root_->numeric_value = value;
    return true;
}

bool cell::bind_value(double value)
{
    root_->type = type::numeric;
    root_->numeric_value = value;
    return true;
}

bool cell::bind_value(const std::string &value)
{
    //Given a value, infer type and display options.
    root_->type = data_type_for_value(value);
    return true;
}

bool cell::bind_value(const char *value)
{
    return bind_value(std::string(value));
}

bool cell::bind_value(bool value)
{
    root_->type = type::boolean;
    root_->bool_value = value;
    return true;
}

bool cell::bind_value(const tm &value)
{
    root_->type = type::date;
    root_->date_value = value;
    return true;
}

coordinate cell::coordinate_from_string(const std::string &coord_string)
{
    // Convert a coordinate string like 'B12' to a tuple ('B', 12)
    bool column_part = true;
    coordinate result;

    for(auto character : coord_string)
    {
        char upper = std::toupper(character, std::locale::classic());

        if(std::isalpha(character, std::locale::classic()))
        {
            if(column_part)
            {
                result.column.append(1, upper);
            }
            else
            {
                std::string msg = "Invalid cell coordinates (" + coord_string + ")";
                throw std::runtime_error(msg);
            }
        }
        else
        {
            if(column_part)
            {
                column_part = false;
            }
            else if(!(std::isdigit(character, std::locale::classic()) || character == '$'))
            {
                std::string msg = "Invalid cell coordinates (" + coord_string + ")";
                throw std::runtime_error(msg);
            }
        }
    }

    std::string row_string = coord_string.substr(result.column.length());
    if(row_string[0] == '$')
    {
        row_string = row_string.substr(1);
    }
    result.row = std::stoi(row_string);

    if(result.row < 1)
    {
        std::string msg = "Invalid cell coordinates (" + coord_string + ")";
        throw std::runtime_error(msg);
    }

    return result;
}

int cell::column_index_from_string(const std::string &column_string)
{
    if(column_string.length() > 3 || column_string.empty())
    {
        throw std::runtime_error("column must be one to three characters");
    }

    int column_index = 0;
    int place = 1;

    for(int i = static_cast<int>(column_string.length()) - 1; i >= 0; i--)
    {
        if(!std::isalpha(column_string[i], std::locale::classic()))
        {
            throw std::runtime_error("column must contain only letters in the range A-Z");
        }

        column_index += (std::toupper(column_string[i], std::locale::classic()) - 'A' + 1) * place;
        place *= 26;
    }

    return column_index;
}

// Convert a column number into a column letter (3 -> 'C')
// Right shift the column col_idx by 26 to find column letters in reverse
// order.These numbers are 1 - based, and can be converted to ASCII
// ordinals by adding 64.
std::string cell::get_column_letter(int column_index)
{
    // these indicies corrospond to A->ZZZ and include all allowed
    // columns
    if(column_index < 1 || column_index > 18278)
    {
	auto msg = "Column index out of bounds: " + std::to_string(column_index);
	throw std::runtime_error(msg);
    }
    
    auto temp = column_index;
    std::string column_letter = "";
    
    while(temp > 0)
    {
	int quotient = temp / 26, remainder = temp % 26;
	
	// check for exact division and borrow if needed
	if(remainder == 0)
	{
	    quotient -= 1;
	    remainder = 26;
	}
	
	column_letter = std::string(1, char(remainder + 64)) + column_letter;
	temp = quotient;
    }
    
    return column_letter;
}

bool cell::is_date() const
{
    return root_->type == type::date;
}

std::string cell::get_address() const
{
    return get_column_letter(root_->column) + std::to_string(root_->row);
}

std::string cell::get_coordinate() const
{
    return get_address();
}

std::string cell::get_hyperlink_rel_id() const
{
    return root_->hyperlink_rel.get_id();
}

bool cell::operator==(std::nullptr_t) const
{
    return root_ == nullptr;
}

bool cell::operator==(const std::string &comparand) const
{
    return root_->type == cell::type::string && root_->string_value == comparand;
}

bool cell::operator==(const char *comparand) const
{
    return *this == std::string(comparand);
}

bool cell::operator==(const tm &comparand) const
{
    return root_->type == cell::type::date && root_->date_value.tm_hour == comparand.tm_hour;
}

bool operator==(const char *comparand, const xlnt::cell &cell)
{
    return cell == comparand;
}

bool operator==(const std::string &comparand, const xlnt::cell &cell)
{
    return cell == comparand;
}

bool operator==(const tm &comparand, const xlnt::cell &cell)
{
    return cell == comparand;
}

style &cell::get_style()
{
    return root_->style;
}

const style &cell::get_style() const
{
    return root_->style;
}

std::string cell::absolute_coordinate(const std::string &absolute_address)
{
    // Convert a coordinate to an absolute coordinate string (B12 -> $B$12)
    auto colon_index = absolute_address.find(':');
    if(colon_index != std::string::npos)
    {
        return absolute_coordinate(absolute_address.substr(0, colon_index)) + ":"
            + absolute_coordinate(absolute_address.substr(colon_index + 1));
    }
    else
    {
        auto coord = coordinate_from_string(absolute_address);
        return std::string("$") + coord.column + "$" + std::to_string(coord.row);
    }
}

xlnt::cell::type cell::get_data_type() const
{
    return root_->type;
}

xlnt::cell cell::get_offset(int column_offset, int row_offset)
{
    return worksheet(root_->parent_worksheet).cell(root_->column + column_offset, root_->row + row_offset);
}

cell &cell::operator=(const cell &rhs)
{
    root_ = rhs.root_;
    return *this;
}

cell &cell::operator=(int value)
{
    root_->type = type::numeric;
    root_->numeric_value = value;
    return *this;
}

cell &cell::operator=(double value)
{
    root_->type = type::numeric;
    root_->numeric_value = value;
    return *this;
}

cell &cell::operator=(bool value)
{
    root_->type = type::boolean;
    root_->bool_value = value;
    return *this;
}

cell &cell::operator=(const std::string &value)
{
    root_->type = data_type_for_value(value);

    switch(root_->type)
    {
    case type::formula:
        root_->formula_value = value;
        break;
    case type::numeric:
        root_->numeric_value = std::stod(value);
        break;
    case type::boolean:
        root_->bool_value = value == "TRUE" || value == "true";
        break;
    case type::error:
        root_->error_value = value;
        break;
    case type::string:
        root_->string_value = value;
        break;
    case type::null:
        break;
    default:
        throw std::runtime_error("bad enum");
    }

    return *this;
}

cell &cell::operator=(const char *value)
{
    return *this = std::string(value);
}

cell &cell::operator=(const tm &value)
{
    root_->type = type::date;
    root_->date_value = value;
    return *this;
}

std::string cell::to_string() const
{
    return root_->to_string();
}

struct worksheet_struct
{
    worksheet_struct(workbook &parent_workbook, const std::string &title)
        : parent_(parent_workbook), title_(title), freeze_panes_(nullptr)
    {

    }

    void garbage_collect()
    {
        for(auto map_iter = cell_map_.begin(); map_iter != cell_map_.end(); map_iter++)
        {
            if(map_iter->second.get_data_type() == cell::type::null)
            {
                map_iter = cell_map_.erase(map_iter);
            }
        }
    }

    std::list<cell> get_cell_collection()
    {
        std::list<xlnt::cell> cells;
        for(auto cell : cell_map_)
        {
            cells.push_front(xlnt::cell(cell.second));
        }
        return cells;
    }

    std::string get_title() const
    {
        return title_;
    }

    void set_title(const std::string &title)
    {
        title_ = title;
    }

    cell get_freeze_panes() const
    {
        return freeze_panes_;
    }

    void set_freeze_panes(cell top_left_cell)
    {
        freeze_panes_ = top_left_cell;
    }

    void set_freeze_panes(const std::string &top_left_coordinate)
    {
        freeze_panes_ = cell(top_left_coordinate);
    }

    void unfreeze_panes()
    {
        freeze_panes_ = xlnt::cell(nullptr);
    }

    xlnt::cell cell(const std::string &coordinate)
    {
        if(cell_map_.find(coordinate) == cell_map_.end())
        {
            auto coord = xlnt::cell::coordinate_from_string(coordinate);
            cell_struct *cell = new xlnt::cell_struct(this, coord.column, coord.row);
            cell_map_[coordinate] = xlnt::cell(cell);
        }

        return cell_map_[coordinate];
    }

    xlnt::cell cell(int row, int column)
    {
        return cell(xlnt::cell::get_column_letter(column + 1) + std::to_string(row + 1));
    }

    int get_highest_row() const
    {
        int highest = 0;
        for(auto cell : cell_map_)
        {
            highest = (std::max)(highest, cell.second.get_row());
        }
        return highest;
    }

    int get_highest_column() const
    {
        int highest = 0;
        for(auto cell : cell_map_)
        {
            highest = (std::max)(highest, xlnt::cell::column_index_from_string(cell.second.get_column()));
        }
        return highest;
    }

    std::string calculate_dimension() const
    {
        int width = get_highest_column();
        std::string width_letter = xlnt::cell::get_column_letter(width);
        int height = get_highest_row();
        return "A1:" + width_letter + std::to_string(height);
    }

    xlnt::range range(const std::string &range_string, int row_offset, int column_offset)
    {
        xlnt::range r;

        auto colon_index = range_string.find(':');

        if(colon_index != std::string::npos)
        {
            auto min_range = range_string.substr(0, colon_index);
            auto max_range = range_string.substr(colon_index + 1);
            auto min_coord = xlnt::cell::coordinate_from_string(min_range);
            auto max_coord = xlnt::cell::coordinate_from_string(max_range);

            if(column_offset != 0)
            {
                min_coord.column = xlnt::cell::get_column_letter(xlnt::cell::column_index_from_string(min_coord.column) + column_offset);
                max_coord.column = xlnt::cell::get_column_letter(xlnt::cell::column_index_from_string(max_coord.column) + column_offset);
            }

            std::unordered_map<int, std::string> column_cache;

            for(int i = xlnt::cell::column_index_from_string(min_coord.column);
                i <= xlnt::cell::column_index_from_string(max_coord.column); i++)
            {
                column_cache[i] = xlnt::cell::get_column_letter(i);
            }
            for(int row = min_coord.row + row_offset; row <= max_coord.row + row_offset; row++)
            {
                r.push_back(std::vector<xlnt::cell>());
                for(int column = xlnt::cell::column_index_from_string(min_coord.column) + column_offset;
                    column <= xlnt::cell::column_index_from_string(max_coord.column) + column_offset; column++)
                {
                    std::string coordinate = column_cache[column] + std::to_string(row);
                    r.back().push_back(cell(coordinate));
                }
            }
        }
        else
        {
            r.push_back(std::vector<xlnt::cell>());
            r.back().push_back(cell(range_string));
        }

        return r;
    }

    relationship create_relationship(const std::string &relationship_type)
    {
        relationships_.push_back(relationship(relationship_type));
        return relationships_.back();
    }

    //void add_chart(chart chart);

    void merge_cells(const std::string &range_string)
    {
        bool first = true;
        for(auto row : range(range_string, 0, 0))
        {
            for(auto cell : row)
            {
                cell.set_merged(true);
                if(!first)
                {
                    cell.bind_value();
                }
                first = false;
            }
        }
    }

    void merge_cells(int start_row, int start_column, int end_row, int end_column)
    {
        auto range_string = xlnt::cell::get_column_letter(start_column + 1) + std::to_string(start_row + 1) + ":"
            + xlnt::cell::get_column_letter(end_column + 1) + std::to_string(end_row + 1);
        merge_cells(range_string);
    }

    void unmerge_cells(const std::string &range_string)
    {
        bool first = true;
        for(auto row : range(range_string, 0, 0))
        {
            for(auto cell : row)
            {
                cell.set_merged(false);
                if(!first)
                {
                    cell.bind_value();
                }
                first = false;
            }
        }
    }

    void unmerge_cells(int start_row, int start_column, int end_row, int end_column)
    {
        auto range_string = xlnt::cell::get_column_letter(start_column + 1) + std::to_string(start_row + 1) + ":"
            + xlnt::cell::get_column_letter(end_column + 1) + std::to_string(end_row + 1);
        merge_cells(range_string);
    }

    void append(const std::vector<std::string> &cells)
    {
        int row = get_highest_row();
        int column = 0;
        for(auto cell : cells)
        {
            this->cell(row, column++) = cell;
        }
    }

    void append(const std::unordered_map<std::string, std::string> &cells)
    {
        int row = get_highest_row();
        for(auto cell : cells)
        {
            int column = xlnt::cell::column_index_from_string(cell.second);
            this->cell(row, column) = cell.second;
        }
    }

    void append(const std::unordered_map<int, std::string> &cells)
    {
        int row = get_highest_row();
        for(auto cell : cells)
        {
            this->cell(row, cell.first) = cell.second;
        }
    }

    xlnt::range rows()
    {
        return range(calculate_dimension(), 0, 0);
    }

    xlnt::range columns()
    {
        throw std::runtime_error("not implemented");
    }

    void operator=(const worksheet_struct &other) = delete;

    workbook &parent_;
    std::string title_;
    xlnt::cell freeze_panes_;
    std::unordered_map<std::string, xlnt::cell> cell_map_;
    std::vector<relationship> relationships_;
    page_setup page_setup_;
};

worksheet::worksheet(worksheet_struct *root) : root_(root)
{
}

worksheet::worksheet(workbook &parent)
{
    *this = parent.create_sheet();
}

page_setup &worksheet::get_page_setup()
{
    return root_->page_setup_;
}

std::string worksheet::to_string() const
{
    return "<Worksheet \"" + root_->title_ + "\">";
}

workbook &worksheet::get_parent() const
{
    return root_->parent_;
}

void worksheet::garbage_collect()
{
    root_->garbage_collect();
}

std::list<cell> worksheet::get_cell_collection()
{
    return root_->get_cell_collection();
}

std::string worksheet::get_title() const
{
    if(root_ == nullptr)
    {
        throw std::runtime_error("null worksheet");
    }
    return root_->title_;
}

void worksheet::set_title(const std::string &title)
{
    root_->title_ = title;
}

cell worksheet::get_freeze_panes() const
{
    return root_->freeze_panes_;
}

void worksheet::set_freeze_panes(xlnt::cell top_left_cell)
{
    root_->set_freeze_panes(top_left_cell);
}

void worksheet::set_freeze_panes(const std::string &top_left_coordinate)
{
    root_->set_freeze_panes(top_left_coordinate);
}

void worksheet::unfreeze_panes()
{
    root_->unfreeze_panes();
}

xlnt::cell worksheet::cell(const std::string &coordinate)
{
    return root_->cell(coordinate);
}

xlnt::cell worksheet::cell(int row, int column)
{
    return root_->cell(row, column);
}

int worksheet::get_highest_row() const
{
    return root_->get_highest_row();
}

int worksheet::get_highest_column() const
{
    return root_->get_highest_column();
}

std::string worksheet::calculate_dimension() const
{
    return root_->calculate_dimension();
}

range worksheet::range(const std::string &range_string, int row_offset, int column_offset)
{
    return root_->range(range_string, row_offset, column_offset);
}

range worksheet::range(const std::string &range_string)
{
    return root_->range(range_string, 0, 0);
}

std::vector<relationship> worksheet::get_relationships()
{
    return root_->relationships_;
}

relationship worksheet::create_relationship(const std::string &relationship_type)
{
    return root_->create_relationship(relationship_type);
}

//void worksheet::add_chart(chart chart);

void worksheet::merge_cells(const std::string &range_string)
{
    root_->merge_cells(range_string);
}

void worksheet::merge_cells(int start_row, int start_column, int end_row, int end_column)
{
    root_->merge_cells(start_row, start_column, end_row, end_column);
}

void worksheet::unmerge_cells(const std::string &range_string)
{
    root_->unmerge_cells(range_string);
}

void worksheet::unmerge_cells(int start_row, int start_column, int end_row, int end_column)
{
    root_->unmerge_cells(start_row, start_column, end_row, end_column);
}

void worksheet::append(const std::vector<std::string> &cells)
{
    root_->append(cells);
}

void worksheet::append(const std::unordered_map<std::string, std::string> &cells)
{
    root_->append(cells);
}

void worksheet::append(const std::unordered_map<int, std::string> &cells)
{
    root_->append(cells);
}

xlnt::range worksheet::rows() const
{
    return root_->rows();
}

xlnt::range worksheet::columns() const
{
    return root_->columns();
}

bool worksheet::operator==(const worksheet &other) const
{
    return root_ == other.root_;
}

bool worksheet::operator!=(const worksheet &other) const
{
    return root_ != other.root_;
}

bool worksheet::operator==(std::nullptr_t) const
{
    return root_ == nullptr;
}

bool worksheet::operator!=(std::nullptr_t) const
{
    return root_ != nullptr;
}


void worksheet::operator=(const worksheet &other)
{
    root_ = other.root_;
}

cell worksheet::operator[](const std::string &address)
{
    return cell(address);
}

std::string writer::write_content_types(workbook &/*wb*/)
{
    /*std::set<std::string> seen;

    pugi::xml_node root;

    if(wb.has_vba_archive())
    {
        root = fromstring(wb.get_vba_archive().read(ARC_CONTENT_TYPES));

        for(auto elem : root.findall("{" + CONTYPES_NS + "}Override"))
        {
            seen.insert(elem.attrib["PartName"]);
        }
    }
    else
    {
        root = Element("{" + CONTYPES_NS + "}Types");

        for(auto content_type : static_content_types_config)
        {
            if(setting_type == "Override")
            {
                tag = "{" + CONTYPES_NS + "}Override";
                attrib = {"PartName": "/" + name};
            }
            else
            {
                tag = "{" + CONTYPES_NS + "}Default";
                attrib = {"Extension": name};
            }

            attrib["ContentType"] = content_type;
            SubElement(root, tag, attrib);
        }
    }

    int drawing_id = 1;
    int chart_id = 1;
    int comments_id = 1;
    int sheet_id = 0;

    for(auto sheet : wb)
    {
        std::string name = "/xl/worksheets/sheet" + std::to_string(sheet_id) + ".xml";

        if(seen.find(name) == seen.end())
        {
            SubElement(root, "{" + CONTYPES_NS + "}Override", {{"PartName", name},
            {"ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"}});
        }

        if(sheet._charts || sheet._images)
        {
            name = "/xl/drawings/drawing" + drawing_id + ".xml";

            if(seen.find(name) == seen.end())
            {
                SubElement(root, "{%s}Override" % CONTYPES_NS, {"PartName" : name,
                    "ContentType" : "application/vnd.openxmlformats-officedocument.drawing+xml"});
            }

            drawing_id += 1;

            for(auto chart : sheet._charts)
            {
                name = "/xl/charts/chart%d.xml" % chart_id;

                if(seen.find(name) == seen.end())
                {
                    SubElement(root, "{%s}Override" % CONTYPES_NS, {"PartName" : name,
                        "ContentType" : "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"});
                }

                chart_id += 1;

                if(chart._shapes)
                {
                    name = "/xl/drawings/drawing%d.xml" % drawing_id;

                    if(seen.find(name) == seen.end())
                    {
                        SubElement(root, "{%s}Override" % CONTYPES_NS, {"PartName" : name,
                            "ContentType" : "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml"});
                    }

                    drawing_id += 1;
                }
            }
        }
        if(sheet.get_comment_count() > 0)
        {
            SubElement(root, "{%s}Override" % CONTYPES_NS,
            {"PartName": "/xl/comments%d.xml" % comments_id,
            "ContentType" : "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"});
            comments_id += 1;
        }
    }*/

    pugi::xml_document doc;
    auto root_node = doc.append_child("Types");
    root_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/package/2006/content-types");

    auto theme_node = root_node.append_child("Override");
    theme_node.append_attribute("PartName").set_value("/xl/theme1.xml");
    theme_node.append_attribute("ContentType").set_value("application/vnd.openxmlformats-officedocument.theme+xml");

    auto styles_node = root_node.append_child("Override");
    styles_node.append_attribute("PartName").set_value("/xl/styles.xml");
    styles_node.append_attribute("ContentType").set_value("application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");

    auto rels_node = root_node.append_child("Default");
    rels_node.append_attribute("Extension").set_value("rels");
    rels_node.append_attribute("ContentType").set_value("application/vnd.openxmlformats-package.relationships+xml");

    auto xml_node = root_node.append_child("Default");
    xml_node.append_attribute("Extension").set_value("xml");
    xml_node.append_attribute("ContentType").set_value("application/xml");

    auto workbook_node = root_node.append_child("Override");
    workbook_node.append_attribute("PartName").set_value("/xl/workbook.xml");
    workbook_node.append_attribute("ContentType").set_value("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");

    auto app_props_node = root_node.append_child("Override");
    app_props_node.append_attribute("PartName").set_value("/docProps/app.xml");
    app_props_node.append_attribute("ContentType").set_value("application/vnd.openxmlformats-officedocument.extended-properties+xml");

    auto sheet_node = root_node.append_child("Override");
    sheet_node.append_attribute("PartName").set_value("/xl/worksheets/sheet1.xml");
    sheet_node.append_attribute("ContentType").set_value("application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");

    auto core_props_node = root_node.append_child("Override");
    core_props_node.append_attribute("PartName").set_value("/docProps/core.xml");
    core_props_node.append_attribute("ContentType").set_value("application/vnd.openxmlformats-package.core-properties+xml");

    std::stringstream ss;
    doc.save(ss);
    return ss.str();
}

std::string writer::write_workbook(workbook &wb)
{
    pugi::xml_document doc;
    auto root_node = doc.append_child("workbook");
    root_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    root_node.append_attribute("xmlns:r").set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    auto sheets_node = root_node.append_child("sheets");
    for(auto ws : wb)
    {
        auto sheet_node = sheets_node.append_child("sheet");
        sheet_node.append_attribute("name").set_value("Sheet1");
        sheet_node.append_attribute("sheetId").set_value("1");
        sheet_node.append_attribute("r:id").set_value("rId1");
    }
    std::stringstream ss;
    doc.save(ss);
    return ss.str();
}

std::string writer::write_root_rels(workbook &/*wb*/)
{
    /*
    if(wb.has_vba_archive())
    {
        // See if there was a customUI relation and reuse its id
        arc = fromstring(workbook.vba_archive.read(ARC_ROOT_RELS));
        rels = arc.findall(relation_tag);
        rId = None;
        for(rel in rels)
        {
            if(rel.get("Target") == ARC_CUSTOM_UI)
            {
                rId = rel.get("Id");
                break;
            }
        }
        if(rId is not None)
        {
            SubElement(root, relation_tag, {"Id": rId, "Target" : ARC_CUSTOM_UI,
                "Type" : "%s" % CUSTOMUI_NS});
        }
    }*/

    pugi::xml_document doc;
    auto root_node = doc.append_child("Relationships");
    root_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/package/2006/relationships");
    auto app_props_node = root_node.append_child("Relationship");
    app_props_node.append_attribute("Id").set_value("rId3");
    app_props_node.append_attribute("Type").set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
    app_props_node.append_attribute("Target").set_value("docProps/app.xml");
    auto core_props_node = root_node.append_child("Relationship");
    core_props_node.append_attribute("Id").set_value("rId3");
    core_props_node.append_attribute("Type").set_value("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
    core_props_node.append_attribute("Target").set_value("docProps/core.xml");
    auto workbook_node = root_node.append_child("Relationship");
    workbook_node.append_attribute("Id").set_value("rId3");
    workbook_node.append_attribute("Type").set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
    workbook_node.append_attribute("Target").set_value("xl/workbook.xml");
    std::stringstream ss;
    doc.save(ss);
    return ss.str();
}

std::string writer::write_worksheet(worksheet ws)
{
    pugi::xml_document doc;
    auto root_node = doc.append_child("worksheet");
    auto title_node = root_node.append_child("title");
    title_node.text().set(ws.get_title().c_str());
    std::stringstream ss;
    doc.save(ss);
    return ss.str();
}

std::string writer::create_temporary_file()
{
    return "a.temp";
}

void writer::delete_temporary_file(const std::string &filename)
{
    std::remove(filename.c_str());
}

void reader::read_workbook(workbook &wb, zip_file &zip)
{
    auto content_types = read_content_types(zip.get_file_contents("[Content_Types].xml"));
    auto root_relationships = read_relationships(zip.get_file_contents("_rels/.rels"));
    auto workbook_relationships = read_relationships(zip.get_file_contents("xl/_rels/workbook.xml.rels"));

    pugi::xml_document doc;
    doc.load(zip.get_file_contents("xl/workbook.xml").c_str());

    auto root_node = doc.child("workbook");
    auto sheets_node = root_node.child("sheets");

    while(!wb.get_sheet_names().empty())
    {
        wb.remove_sheet(wb.get_sheet_by_name(wb.get_sheet_names().front()));
    }

    for(auto sheet_node : sheets_node.children("sheet"))
    {
        auto relation_id = sheet_node.attribute("r:id").as_string();
        auto ws = wb.create_sheet(sheet_node.attribute("name").as_string());
        std::string sheet_filename("xl/");
        sheet_filename += workbook_relationships[relation_id].second;
        read_worksheet(ws, zip.get_file_contents(sheet_filename));
    }
}

void reader::read_worksheet(worksheet ws, const std::string &content)
{
    pugi::xml_document doc;
    doc.load(content.c_str());

    auto root_node = doc.child("worksheet");

    ws.get_page_setup().orientation = xlnt::page_setup::Orientation::Landscape;
}

std::unordered_map<std::string, std::pair<std::string, std::string>> reader::read_relationships(const std::string &content)
{
    pugi::xml_document doc;
    doc.load(content.c_str());

    auto root_node = doc.child("Relationships");

    std::unordered_map<std::string, std::pair<std::string, std::string>> relationships;

    for(auto relationship : root_node.children("Relationship"))
    {
        std::string id = relationship.attribute("Id").as_string();
        std::string type = relationship.attribute("Type").as_string();
        std::string target = relationship.attribute("Target").as_string();
        relationships[id] = std::make_pair(type, target);
    }

    return relationships;
}

std::pair<std::unordered_map<std::string, std::string>, std::unordered_map<std::string, std::string>> reader::read_content_types(const std::string &content)
{
    pugi::xml_document doc;
    doc.load(content.c_str());

    auto root_node = doc.child("Types");

    std::unordered_map<std::string, std::string> default_types;

    for(auto child : root_node.children("Default"))
    {
        default_types[child.attribute("Extension").as_string()] = child.attribute("ContentType").as_string();
    }

    std::unordered_map<std::string, std::string> override_types;

    for(auto child : root_node.children("Override"))
    {
        override_types[child.attribute("PartName").as_string()] = child.attribute("ContentType").as_string();
    }

    return std::make_pair(default_types, override_types);
}

workbook::workbook(optimized optimized) 
: active_sheet_index_(0), 
optimized_read_(optimized==optimized::read), 
optimized_write_(optimized==optimized::write)
{
    if(!optimized_write_)
    {
        auto ws = create_sheet();
        ws.set_title("Sheet1");
    }
}

worksheet workbook::get_sheet_by_name(const std::string &name)
{
    auto match = std::find_if(worksheets_.begin(), worksheets_.end(), [&](const worksheet &w) { return w.get_title() == name; });
    if(match != worksheets_.end())
    {
        return worksheet(*match);
    }
    return worksheet(nullptr);
}

worksheet workbook::get_active_sheet()
{
    return worksheets_[active_sheet_index_];
}

worksheet workbook::create_sheet()
{
    std::string title = "Sheet1";
    int index = 1;
    while(get_sheet_by_name(title) != nullptr)
    {
        title = "Sheet" + std::to_string(++index);
    }
    auto *worksheet = new worksheet_struct(*this, title);
    worksheets_.push_back(worksheet);
    return get_sheet_by_name(title);
}

void workbook::create_named_range(const std::string &/*range_string*/, worksheet /*ws*/, const std::string &/*name*/)
{

}

void workbook::load(const std::string &filename)
{
    zip_file f(filename, file_mode::open);
    reader::read_workbook(*this, f);
}

void workbook::remove_sheet(worksheet ws)
{
    auto match_iter = std::find(worksheets_.begin(), worksheets_.end(), ws);
    if(match_iter == worksheets_.end())
    {
        throw std::runtime_error("worksheet not owned by this workbook");
    }
    delete match_iter->root_;
    worksheets_.erase(match_iter);
}

worksheet workbook::create_sheet(std::size_t index)
{
    auto ws = create_sheet();
    if(index != worksheets_.size())
    {
        std::swap(worksheets_[index], worksheets_.back());
    }
    return ws;
}

worksheet workbook::create_sheet(const std::string &title)
{
    if(title.length() > 31)
    {
        throw bad_sheet_title(title);
    }

    if(std::find_if(title.begin(), title.end(), 
        [](char c) { return !(std::isalpha(c, std::locale::classic()) 
        || std::isdigit(c, std::locale::classic())); }) != title.end())
    {
        throw bad_sheet_title(title);
    }

    if(get_sheet_by_name(title) != nullptr)
    {
        throw std::runtime_error("sheet exists");
    }
    auto *worksheet = new worksheet_struct(*this, title);
    worksheets_.push_back(worksheet);
    return get_sheet_by_name(title);
}

workbook &workbook::optimized_write(bool optimized_write)
{
    optimized_write_ = optimized_write;
    return *this;
}

std::vector<worksheet>::iterator workbook::begin()
{
    return worksheets_.begin();
}

std::vector<worksheet>::iterator workbook::end()
{
    return worksheets_.end();
}

std::vector<std::string> workbook::get_sheet_names() const
{
    std::vector<std::string> names;
    for(auto &ws : worksheets_)
    {
        names.push_back(ws.get_title());
    }
    return names;
}

worksheet workbook::operator[](const std::string &name)
{
    return get_sheet_by_name(name);
}

worksheet workbook::operator[](int index)
{
    return worksheets_[index];
}


void workbook::save(const std::string &filename)
{
    zip_file f(filename, file_mode::create, file_access::write);
    f.set_file_contents("[Content_Types].xml", writer::write_content_types(*this));
    f.set_file_contents("rels/.rels", writer::write_root_rels(*this));
    f.set_file_contents("xl/workbook.xml", writer::write_workbook(*this));
    for(worksheet ws : *this)
    {
        f.set_file_contents("xl/worksheets/sheet.xml", writer::write_worksheet(ws));
    }
}

std::string cell_struct::to_string() const
{
    return "<Cell " + parent_worksheet->title_ + "." + xlnt::cell::get_column_letter(column + 1) + std::to_string(row) + ">";
}

bool workbook::operator==(const workbook &rhs) const
{
    if(optimized_write_ == rhs.optimized_write_
        && optimized_read_ == rhs.optimized_read_
        && guess_types_ == rhs.guess_types_
        && data_only_ == rhs.data_only_
        && active_sheet_index_ == rhs.active_sheet_index_
        && encoding_ == rhs.encoding_)
    {
        if(worksheets_.size() != rhs.worksheets_.size())
        {
            return false;
        }

        for(int i = 0; i < worksheets_.size(); i++)
        {
            if(worksheets_[i] != rhs.worksheets_[i])
            {
                return false;
            }
        }

        /*
        if(named_ranges_.size() != rhs.named_ranges_.size())
        {
            return false;
        }

        for(int i = 0; i < named_ranges_.size(); i++)
        {
            if(named_ranges_[i] != rhs.named_ranges_[i])
            {
                return false;
            }
        }

        if(relationships_.size() != rhs.relationships_.size())
        {
            return false;
        }

        for(int i = 0; i < relationships_.size(); i++)
        {
            if(relationships_[i] != rhs.relationships_[i])
            {
                return false;
            }
        }

        if(drawings_.size() != rhs.drawings_.size())
        {
            return false;
        }

        for(int i = 0; i < drawings_.size(); i++)
        {
            if(drawings_[i] != rhs.drawings_[i])
            {
                return false;
            }
        }
        */

        return true;
    }
    return false;
}

void sheet_protection::set_password(const std::string &password)
{
    hashed_password_ = hash_password(password);
}

std::string sheet_protection::hash_password(const std::string &password)
{
    return password;
}

int string_table::operator[](const std::string &/*key*/) const
{
    return 0;
}

void string_table_builder::add(const std::string &/*string*/)
{

}

}
