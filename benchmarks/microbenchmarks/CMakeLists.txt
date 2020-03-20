# FetchContent added in cmake v3.11
# https://cmake.org/cmake/help/v3.11/module/FetchContent.html
# this file is behind a feature flag (XLNT_MICROBENCH_ENABLED) so the primary build is not affected
cmake_minimum_required(VERSION 3.11)
project(xlnt_ubench)

# acquire google benchmark dependency
# disable generation of the various test projects
set(BENCHMARK_ENABLE_TESTING OFF)
# gtest not required
set(BENCHMARK_ENABLE_GTEST_TESTS OFF)

include(FetchContent)
FetchContent_Declare(
	googlebenchmark
	GIT_REPOSITORY 	https://github.com/google/benchmark
	GIT_TAG			v1.5.0
)
# download if not already present
FetchContent_GetProperties(googlebenchmark)
if(NOT googlebenchmark_POPULATED)
	FetchContent_Populate(googlebenchmark)
	add_subdirectory(${googlebenchmark_SOURCE_DIR} ${googlebenchmark_BINARY_DIR})
endif()
# equivalent of add_subdirectory, now available for use
FetchContent_MakeAvailable(googlebenchmark)


add_executable(xlnt_ubench)
target_sources(xlnt_ubench
	PRIVATE
		string_to_double.cpp
		double_to_string.cpp
)
target_link_libraries(xlnt_ubench benchmark_main xlnt)
target_compile_features(xlnt_ubench PRIVATE cxx_std_17)