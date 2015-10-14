cmake_minimum_required(VERSION 2.8.9)
project(xlnt)

set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/../../../lib)
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/../../../bin)

add_subdirectory(xlnt)
add_subdirectory(xlnt.test)
