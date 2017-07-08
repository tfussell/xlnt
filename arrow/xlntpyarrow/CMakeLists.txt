cmake_minimum_required(VERSION 3.2)
project(xlntpyarrow)

if(NOT MSVC)
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -std=c++11")
endif()

configure_file("${CMAKE_CURRENT_SOURCE_DIR}/setup.py.cmake"
    "${CMAKE_CURRENT_BINARY_DIR}/setup.py")

add_custom_target(xlntpyarrow ALL
    COMMAND python setup.py build
    DEPENDS xlntarrow)
