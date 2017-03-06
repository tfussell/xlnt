cmake_minimum_required(VERSION 3.2)
project(program)

set(CMAKE_MODULE_PATH ${CMAKE_MODULE_PATH} "${CMAKE_CURRENT_SOURCE_DIR}")
add_subdirectory(${CMAKE_CURRENT_SOURCE_DIR}/libstudxml)

if(BUILD_DEPS)
    add_subdirectory(${CMAKE_CURRENT_SOURCE_DIR}/zlib)
    add_subdirectory(${CMAKE_CURRENT_SOURCE_DIR}/cryptopp)
    add_dependencies(libstudxml zlib-download cryptopp-download)
endif()
