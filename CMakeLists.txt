cmake_minimum_required(VERSION 3.2)
project(xlnt_all)

# This indicates to CMakeLists in subdirectories that they are part of a larger project
set(COMBINED_PROJECT TRUE)

# Library type
option(STATIC "Set to ON to build xlnt as a static library instead of a shared library" OFF)

# Optional components
option(SAMPLES "Set to ON to build executable code samples (in ./samples)" OFF)
option(BENCHMARKS "Set to ON to build performance benchmarks (in ./benchmarks)" OFF)

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

find_package(PythonInterp)
find_package(CxxTest)

if(PYTHONINTERP_FOUND AND CXXTEST_FOUND)
    add_subdirectory(tests)
elseif(NOT PYTHONINTERP_FOUND)
    message("Python couldn't be found in the current PATH but is required for building tests. Tests will be skipped for now.")
else()
    message("CxxTest couldn't be found in the current PATH but is required for building tests. Tests will be skipped for now.")
endif()

add_subdirectory(source)
