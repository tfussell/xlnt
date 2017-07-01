cmake_minimum_required(VERSION 3.2)
project(xlnt_arrow)

set(CMAKE_CXX_STANDARD 14)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(XLNT_ARROW
    ${CMAKE_CURRENT_SOURCE_DIR}/xlnt_arrow.hpp
    ${CMAKE_CURRENT_SOURCE_DIR}/xlnt_arrow.cpp)

add_library(xlnt_arrow STATIC ${XLNT_ARROW})
target_include_directories(xlnt_arrow PRIVATE ${CMAKE_CURRENT_SOURCE_DIR})
target_include_directories(xlnt_arrow PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/../../third-party/arrow/cpp/src)
target_include_directories(xlnt_arrow PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/../../include)
