cmake_minimum_required(VERSION 3.1)
project(xlnt_all)

# This indicates to CMakeLists in subdirectories that they are part of a larger project
set(COMBINED_PROJECT TRUE)

# Library type
option(STATIC "Set to ON to build ${PROJECT_NAME} as a static library instead of a shared library" OFF)

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

if(PYTHONINTERP_FOUND)
    add_subdirectory(tests)
else()
    message("Python couldn't be found in the current PATH but is required for building tests. Tests will be skipped for now.")
endif()

add_subdirectory(source)
