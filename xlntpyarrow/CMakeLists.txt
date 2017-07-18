cmake_minimum_required(VERSION 3.2)
project(xlntpyarrow)

if(NOT MSVC)
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -std=c++11")
endif()

option(CONDA_ROOT "Path to Conda directory." "")

if(NOT CONDA_ROOT)
	message(FATAL_ERROR "Missing Conda root directory option (-D CONDA_ROOT).")
endif()

# Replace backslash with forward slash since Python interprets it as an escaped character
string(REGEX REPLACE "\\\\" "/" CONDA_ROOT ${CONDA_ROOT})

configure_file("${CMAKE_CURRENT_SOURCE_DIR}/setup.py.cmake"
	 "${CMAKE_CURRENT_BINARY_DIR}/setup.py")

add_custom_target(xlntpyarrow ALL
    COMMAND python setup.py install --record=record.txt --xlntlib=$<TARGET_FILE:xlnt>
    DEPENDS xlnt
    SOURCES xlntpyarrow.cpp setup.py.cmake methods.hpp methods.cpp)
