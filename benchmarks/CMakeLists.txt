cmake_minimum_required(VERSION 3.2)
project(xlnt.benchmarks)

set(CMAKE_CXX_STANDARD 14)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

if(NOT COMBINED_PROJECT)
    add_subdirectory(${CMAKE_CURRENT_SOURCE_DIR}/../source ${CMAKE_CURRENT_BINARY_DIR}/source)
endif()

set(XLNT_BENCHMARK_DATA_DIR ${CMAKE_CURRENT_SOURCE_DIR}/data)

file(GLOB BENCHMARK_SOURCES ${CMAKE_CURRENT_SOURCE_DIR}/*.cpp)

foreach(BENCHMARK_SOURCE IN ITEMS ${BENCHMARK_SOURCES})
    get_filename_component(BENCHMARK_NAME ${BENCHMARK_SOURCE} NAME_WE)
    set(BENCHMARK_EXECUTABLE benchmark-${BENCHMARK_NAME})

    add_executable(${BENCHMARK_EXECUTABLE} ${BENCHMARK_SOURCE})

    target_link_libraries(${BENCHMARK_EXECUTABLE} PRIVATE xlnt)
	target_include_directories(${BENCHMARK_EXECUTABLE}
		PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/../tests)
	target_compile_definitions(${BENCHMARK_EXECUTABLE} PRIVATE XLNT_BENCHMARK_DATA_DIR=${XLNT_BENCHMARK_DATA_DIR})

    if(MSVC AND NOT STATIC)
        add_custom_command(TARGET ${BENCHMARK_EXECUTABLE} POST_BUILD 
            COMMAND ${CMAKE_COMMAND} -E copy_if_different
            $<TARGET_FILE:xlnt>
            $<TARGET_FILE_DIR:${BENCHMARK_EXECUTABLE}>)
    endif()    
endforeach()
