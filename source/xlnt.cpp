#include <algorithm>
#include <array>
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

struct part_struct
{
    part_struct(package_impl &package, const std::string &uri_part, const std::string &mime_type = "", compression_option compression = compression_option::NotCompressed)
        : compression_option_(compression),
	  content_type_(mime_type),
	  package_(package),
	  uri_(uri_part)
    {}

    part_struct(package_impl &package, const std::string &uri, opcContainer *container)
        : package_(package),
        uri_(uri),
        container_(container)
    {
    }

    relationship create_relationship(const std::string &target_uri, target_mode target_mode, const std::string &relationship_type);

    void delete_relationship(const std::string &id);

    relationship get_relationship(const std::string &id);

    relationship_collection get_relationships();

    relationship_collection get_relationship_by_type(const std::string &relationship_type);

    std::string read()
    {
        std::string ss;
        auto part_stream = opcContainerOpenInputStream(container_, (xmlChar*)uri_.c_str());
        std::array<xmlChar, 1024> buffer;
        auto bytes_read = opcContainerReadInputStream(part_stream, buffer.data(), static_cast<opc_uint32_t>(buffer.size()));
        if(bytes_read > 0)
        {
            ss.append(std::string(buffer.begin(), buffer.begin() + bytes_read));
            while(bytes_read == buffer.size())
            {
                auto bytes_read = opcContainerReadInputStream(part_stream, buffer.data(), static_cast<opc_uint32_t>(buffer.size()));
                ss.append(std::string(buffer.begin(), buffer.begin() + bytes_read));
            }
        }
        opcContainerCloseInputStream(part_stream);
        return ss;
    }

    void write(const std::string &data)
    {
        auto name = uri_;
        auto name_pointer = name.c_str();

        auto match = opcPartFind(container_, (xmlChar*)name_pointer, nullptr, 0);
        if(match == nullptr)
        {
            match = opcPartCreate(container_, (xmlChar*)name_pointer, nullptr, 0);
        }

        auto part_stream = opcContainerCreateOutputStream(container_, (xmlChar*)name_pointer, opcCompressionOption_t::OPC_COMPRESSIONOPTION_NORMAL);

        std::stringstream ss(data);
        std::array<xmlChar, 1024> buffer;

        while(ss)
        {
            ss.get((char*)buffer.data(), 1024);
            auto count = ss.gcount();
            if(count > 0)
            {
                opcContainerWriteOutputStream(part_stream, buffer.data(), static_cast<opc_uint32_t>(count));
            }
        }
        opcContainerCloseOutputStream(part_stream);
    }

    /// <summary>
    /// Returns a value that indicates whether this part owns a relationship with a specified Id.
    /// </summary>
    bool relationship_exists(const std::string &id) const;

    /// <summary>
    /// Returns true if the given Id string is a valid relationship identifier.
    /// </summary>
    bool is_valid_xml_id(const std::string &id);

    void operator=(const part_struct &other);

    compression_option compression_option_;
    std::string content_type_;
    package_impl &package_;
    std::string uri_;
    opcContainer *container_;
};

part::part(part_struct *root) : root_(root)
{

}

std::string part::get_content_type() const
{
    return "";
}

std::string part::read()
{
    if(root_ == nullptr)
    {
        return "";
    }

    return root_->read();
}

void part::write(const std::string &data)
{
    if(root_ == nullptr)
    {
        return;
    }

    return root_->write(data);
}

bool part::operator==(const part &comparand) const
{
    return root_ == comparand.root_;
}

bool part::operator==(const std::nullptr_t &) const
{
    return root_ == nullptr;
}

struct package_impl
{
    static const int BufferSize = 4096 * 4;

    package *parent_;
    opcContainer *opc_container_;
    std::iostream &stream_;
    std::fstream file_stream_;
    file_mode package_mode_;
    file_access package_access_;
    std::vector<xmlChar> container_buffer_;
    bool open_;
    bool streaming_;

    file_access get_file_open_access() const
    {
        if(!open_)
        {
            throw std::runtime_error("The package is not open");
        }

        return package_access_;
    }

    package_impl(std::iostream &stream, file_mode package_mode, file_access package_access)
        : stream_(stream), 
	  package_mode_(package_mode), 
	  package_access_(package_access),
	  container_buffer_(BufferSize)
    {
    }

    package_impl(const std::string &path, file_mode package_mode, file_access package_access)
        : stream_(file_stream_),
        package_mode_(package_mode),
        package_access_(package_access),
        container_buffer_(BufferSize),
        filename_(path)
    {
        switch(package_mode)
        {
        case file_mode::Append:
            switch(package_access)
            {
            case file_access::Read: throw std::runtime_error("Append can only be used with file_access.Write");
            case file_access::ReadWrite: throw std::runtime_error("Append can only be used with file_access.Write");
            case file_access::Write:
                file_stream_.open(path, std::ios::binary | std::ios::app | std::ios::out);
                break;
            default: throw std::runtime_error("invalid access");
            }
            break;
        case file_mode::Create:
            switch(package_access)
            {
            case file_access::Read:
                file_stream_.open(path, std::ios::binary | std::ios::in | std::ios::trunc);
                break;
            case file_access::ReadWrite:
                file_stream_.open(path, std::ios::binary | std::ios::in | std::ios::out | std::ios::trunc);
                break;
            case file_access::Write:
                file_stream_.open(path, std::ios::binary | std::ios::out);
                break;
            default: throw std::runtime_error("invalid access");
            }
            break;
        case file_mode::CreateNew:
            if(!file::exists(path))
            {
                throw std::runtime_error("File already exists");
            }
            switch(package_access)
            {
            case file_access::Read:
                file_stream_.open(path, std::ios::binary | std::ios::in);
                break;
            case file_access::ReadWrite:
                file_stream_.open(path, std::ios::binary | std::ios::in | std::ios::out);
                break;
            case file_access::Write:
                file_stream_.open(path, std::ios::binary | std::ios::out);
                break;
            default: throw std::runtime_error("invalid access");
            }
            break;
        case file_mode::Open:
            if(!file::exists(path))
            {
                throw std::runtime_error("Can't open non-existent file");
            }
            switch(package_access)
            {
            case file_access::Read:
                file_stream_.open(path, std::ios::binary | std::ios::in);
                break;
            case file_access::ReadWrite:
                file_stream_.open(path, std::ios::binary | std::ios::in | std::ios::out);
                break;
            case file_access::Write:
                file_stream_.open(path, std::ios::binary | std::ios::out);
                break;
            default: throw std::runtime_error("invalid access");
            }
            break;
        case file_mode::OpenOrCreate:
            switch(package_access)
            {
            case file_access::Read:
                file_stream_.open(path, std::ios::binary | std::ios::in);
                break;
            case file_access::ReadWrite:
                file_stream_.open(path, std::ios::binary | std::ios::in | std::ios::out);
                break;
            case file_access::Write:
                file_stream_.open(path, std::ios::binary | std::ios::out);
                break;
            default: throw std::runtime_error("invalid access");
            }
            break;
        case file_mode::Truncate:
            if(!file::exists(path))
            {
                throw std::runtime_error("Can't truncate non-existent file");
            }
            switch(package_access)
            {
            case file_access::Read:
                file_stream_.open(path, std::ios::binary | std::ios::trunc | std::ios::in);
                break;
            case file_access::ReadWrite:
                file_stream_.open(path, std::ios::binary | std::ios::trunc | std::ios::in | std::ios::out);
                break;
            case file_access::Write:
                file_stream_.open(path, std::ios::binary | std::ios::trunc | std::ios::out);
                break;
            default: throw std::runtime_error("invalid access");
            }
            break;
        }

        if(!file_stream_)
        {
            throw std::runtime_error("something");
        }
    }

    void open_container()
    {
        if(package_mode_ != file_mode::Open)
        {
            stream_.write((const char *)existing_xlsx.data(), existing_xlsx.size());
            stream_.seekg(std::ios::beg);
            stream_.seekp(std::ios::beg);
            stream_.flush();
        }

        opcContainerOpenMode m;

        switch(package_access_)
        {
        case file_access::Read:
            m = opcContainerOpenMode::OPC_OPEN_READ_ONLY;
            break;
        case file_access::ReadWrite:
            m = opcContainerOpenMode::OPC_OPEN_READ_WRITE;
            break;
        case file_access::Write:
            m = opcContainerOpenMode::OPC_OPEN_WRITE_ONLY;
            break;
        default:
            throw std::runtime_error("unknown file access");
        }

        opc_container_ = opcContainerOpenIO(&read_callback, &write_callback,
            &close_callback, &seek_callback,
            &trim_callback, &flush_callback, this, BufferSize, m, this);
        open_ = true;

        type_ = determine_type();

        if(type_ != package::type::Excel)
        {
            throw std::runtime_error("only excel spreadsheets are supported for now");
        }
    }

    ~package_impl()
    {
        close();
    }

    package::type determine_type()
    {
        opcRelation rel = opcRelationFind(opc_container_, OPC_PART_INVALID, NULL, _X("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"));

        if(OPC_RELATION_INVALID != rel)
        {
            opcPart main = opcRelationGetInternalTarget(opc_container_, OPC_PART_INVALID, rel);

            if(OPC_PART_INVALID != main)
            {
                const xmlChar *type = opcPartGetType(opc_container_, main);

                if(0 == xmlStrcmp(type, _X("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")))
                {
                    return package::type::Word;
                }
                else if(0 == xmlStrcmp(type, _X("application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")))
                {
                    return package::type::Powerpoint;
                }
                else if(0 == xmlStrcmp(type, _X("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")))
                {
                    return package::type::Excel;
                }
            }
        }

        return package::type::Zip;
    }

    void close()
    {
        if(open_)
        {
            open_ = false;
            for(auto part : parts_)
            {
                delete part.second;
            }
            opcContainerClose(opc_container_, opcContainerCloseMode::OPC_CLOSE_NOW);
            opc_container_ = nullptr;
        }
    }

    part create_part(const std::string &part_uri, const std::string &/*content_type*/, compression_option /*compression*/)
    {
        if(parts_.find(part_uri) == parts_.end())
        {
            parts_[part_uri] = new part_struct(*this, part_uri, opc_container_);
        }
        return part(parts_[part_uri]);
    }

    void delete_part(const std::string &/*part_uri*/)
    {

    }

    void flush()
    {
        stream_.flush();
    }

    part get_part(const std::string &part_uri)
    {
        if(parts_.find(part_uri) == parts_.end())
        {
            parts_[part_uri] = new part_struct(*this, part_uri, opc_container_);
        }
        return part(parts_[part_uri]);
    }

    part_collection get_parts()
    {
        return part_collection();
    }

    int write(char *buffer, int length)
    {
        stream_.read(buffer, length);
        auto bytes_read = stream_.gcount();
        return static_cast<int>(bytes_read);
    }

    int read(const char *buffer, int length)
    {
        auto before = stream_.tellp();
        stream_.write(buffer, length);
        auto bytes_written = stream_.tellp() - before;
        return static_cast<int>(bytes_written);
    }

    std::ios::pos_type seek(std::ios::pos_type offset)
    {
        stream_.clear();
        stream_.seekg(offset);
        auto current_position = stream_.tellg();
        if(stream_.eof())
        {
            std::cout << "eof" << std::endl;
        }
        if(stream_.fail())
        {
            std::cout << "fail" << std::endl;
        }
        if(stream_.bad())
        {
            std::cout << "bad" << std::endl;
        }
        return current_position;
    }

    int trim(std::ios::pos_type new_size)
    {
        file_stream_.flush();
        std::vector<char> buffer(new_size);
        auto current_position = stream_.tellg();
        seek(std::ios::beg);
        write(buffer.data(), (int)new_size);
        file_stream_.close();
        file_stream_.open(filename_, std::ios::trunc | std::ios::out | std::ios::binary | std::ios::in);
        file_stream_.write(buffer.data(), new_size);
        seek(current_position);

        return 0;
    }

    static int read_callback(void *context, char *buffer, int length)
    {
        auto object = static_cast<package_impl *>(context);
        return object->write(buffer, length);
    }

    static int write_callback(void *context, const char *buffer, int length)
    {
        auto object = static_cast<package_impl *>(context);
        return object->read(buffer, length);
    }

    static int close_callback(void *context)
    {
        auto object = static_cast<package_impl *>(context);
        object->close();
        return 0;
    }

    static opc_ofs_t seek_callback(void *context, opc_ofs_t ofs)
    {
        auto object = static_cast<package_impl *>(context);
        return static_cast<opc_ofs_t>(object->seek(ofs));
    }

    static int trim_callback(void *context, opc_ofs_t new_size)
    {
        auto object = static_cast<package_impl *>(context);
        return object->trim(new_size);
    }

    static int flush_callback(void *context)
    {
        auto object = static_cast<package_impl *>(context);
        object->flush();
        return 0;
    }

    std::unordered_map<std::string, part_struct *> parts_;
    package::type type_;
    std::string filename_;
};

file_access package::get_file_open_access() const
{
    return impl_->get_file_open_access();
}

package::~package()
{
    close();
    delete impl_;
}

void package::open(std::iostream &stream, file_mode package_mode, file_access package_access)
{
    if(impl_ != nullptr)
    {
        close();
        delete impl_;
        impl_ = nullptr;
    }

    impl_ = new package_impl(stream, package_mode, package_access);
    impl_->parent_ = this;
    open_container();
}

void package::open(const std::string &path, file_mode package_mode, file_access package_access, file_share /*package_share*/)
{
    if(impl_ != nullptr)
    {
        close();
        delete impl_;
        impl_ = nullptr;
    }

    impl_ = new package_impl(path, package_mode, package_access);
    impl_->parent_ = this;
    open_container();
}

package::package() : impl_(nullptr)
{
}

void package::open_container()
{
    impl_->open_container();
}

void package::close()
{
    impl_->close();
}

part package::create_part(const std::string &part_uri, const std::string &content_type, compression_option compression)
{
    return impl_->create_part(part_uri, content_type, compression);
}

void package::delete_part(const std::string &part_uri)
{
    impl_->delete_part(part_uri);
}

void package::flush()
{
    impl_->flush();
}

part package::get_part(const std::string &part_uri)
{
    return impl_->get_part(part_uri);
}

part_collection package::get_parts()
{
    return impl_->get_parts();
}

int package::write(char *buffer, int length)
{
    return impl_->write(buffer, length);
}

int package::read(const char *buffer, int length)
{
    return impl_->read(buffer, length);
}

std::ios::pos_type package::seek(std::ios::pos_type offset)
{
    return impl_->seek(offset);
}

int package::trim(std::ios::pos_type new_size)
{
    return impl_->trim(new_size);
}

bool package::operator==(const package &comparand) const
{
    return impl_ == comparand.impl_;
}

bool package::operator==(const std::nullptr_t &) const
{
    return impl_ == nullptr;
}


struct cell_struct
{
    cell_struct(worksheet_struct *ws, const std::string &column, int row)
        : type(cell::type::null), parent_worksheet(ws),
        column(xlnt::cell::column_index_from_string(column) - 1), row(row)
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

cell::type cell::data_type_for_value(const std::string &value)
{
    if(value[0] == '=')
    {
        return type::formula;
    }

    return type::null;
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

bool operator==(const char *comparand, const cell &cell)
{
    return cell == comparand;
}

bool operator==(const std::string &comparand, const cell &cell)
{
    return cell == comparand;
}

bool operator==(const tm &comparand, const cell &cell)
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

cell::type cell::get_data_type() const
{
    return root_->type;
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
    if(value == "")
    {
	root_->type = type::null;
	return *this;
    }

    if(value[0] == '=')
    {
	root_->type = type::formula;
	root_->formula_value = value;
	return *this;
    }

    if(value[0] == '0')
    {
	if(value.length() > 1)
	{
	    if(value[1] == '.')
	    {
		try
		{
		    double double_value = std::stod(value);
		    *this = double_value;
		    return *this;
		}
		catch(std::invalid_argument)
		{
		}
	    }
	}
	else
	{
	    root_->type = type::numeric;
	    root_->numeric_value = 0;
	    return *this;
	}

	root_->type = type::string;
	root_->string_value = value;
	return *this;
    }

    try
    {
	double double_value = std::stod(value);
	*this = double_value;
    }
    catch(std::invalid_argument)
    {
	root_->type = type::string;
	root_->string_value = value;
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
        for(auto cell : cells)
        {
            cell_map_["a"] = cell;
        }
    }

    void append(const std::unordered_map<std::string, std::string> &cells)
    {
        for(auto cell : cells)
        {
            cell_map_[cell.first] = cell.second;
        }
    }

    void append(const std::unordered_map<int, std::string> &cells)
    {
        for(auto cell : cells)
        {
            cell_map_[xlnt::cell::get_column_letter(cell.first + 1)] = cell.second;
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
};

worksheet::worksheet(worksheet_struct *root) : root_(root)
{
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

std::string writer::write_content_types(workbook &wb)
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
    }

    return get_document_content(root);*/
    return "";
}

std::string writer::write_root_rels(workbook &wb)
{
    /*root = Element("{%s}Relationships" % PKG_REL_NS);
    relation_tag = "{%s}Relationship" % PKG_REL_NS;
    SubElement(root, relation_tag, {"Id": "rId1", "Target" : ARC_WORKBOOK,
        "Type" : "%s/officeDocument" % REL_NS});
    SubElement(root, relation_tag, {"Id": "rId2", "Target" : ARC_CORE,
        "Type" : "%s/metadata/core-properties" % PKG_REL_NS});
    SubElement(root, relation_tag, {"Id": "rId3", "Target" : ARC_APP,
        "Type" : "%s/extended-properties" % REL_NS});
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
    }

    return get_document_content(root);*/
    return "";
}

workbook::workbook() : active_sheet_index_(0)
{
    auto ws = create_sheet();
    ws.set_title("Sheet1");
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

worksheet workbook::create_sheet(std::size_t index)
{
    auto ws = create_sheet();
    if(index != worksheets_.size())
    {
        std::swap(worksheets_[index], worksheets_.back());
    }
    return ws;
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

void workbook::save(const std::string &filename)
{
    package p;
    p.open(filename);
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

}
