cmake_minimum_required(VERSION 3.2)
project(xlnt_all)

# This indicates to CMakeLists in subdirectories that they are part of a larger project
set(COMBINED_PROJECT TRUE)

# Library type
option(STATIC "Set to ON to build xlnt as a static library instead of a shared library" OFF)

# Optional components
option(TESTS "Set to OFF to skip building test executable (in ./tests)" ON)
option(SAMPLES "Set to ON to build executable code samples (in ./samples)" OFF)
option(BENCHMARKS "Set to ON to build performance benchmarks (in ./benchmarks)" OFF)
option(ARROW "Set to ON to build Arrow conversion functions (in ./contrib/xlntarrow)" OFF)

# Platform specific options
if(NOT MSVC)
    option(COVERAGE "Generate coverage data using gcov" OFF)
endif()

# Add components according to selected options
if(SAMPLES)
    add_subdirectory(samples)
endif()

if(BENCHMARKS)
    add_subdirectory(benchmarks)
endif()

if(TESTS)
    add_subdirectory(tests)
endif()

if(PYTHON)
    add_subdirectory(python)
endif()

add_subdirectory(source)
