cmake_minimum_required(VERSION 3.1)
set(LIBRARY_NAME xlnt)
project(${LIBRARY_NAME}_all VERSION ${LIBRARY_VERSION} LANGUAGES CXX C)

# This indicates to CMakeLists in subdirectories that they are part of a larger project
set(COMBINED_PROJECT 1)

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

add_subdirectory(tests)
add_subdirectory(source)
