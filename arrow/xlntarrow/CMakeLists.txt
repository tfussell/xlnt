cmake_minimum_required(VERSION 3.2)
project(xlntarrow)

set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

option(CONDA_ROOT "Path to Conda directory." "")
if(NOT CONDA_ROOT)
    message(FATAL_ERROR "Missing Conda root directory option (-D CONDA_ROOT).")
endif()

set(XLNT_ARROW
    ${CMAKE_CURRENT_SOURCE_DIR}/xlntarrow.hpp
    ${CMAKE_CURRENT_SOURCE_DIR}/xlntarrow.cpp)

link_directories(${CONDA_ROOT}/lib)

if(NOT STATIC)
	add_library(xlntarrow SHARED ${XLNT_ARROW})
	target_compile_definitions(xlntarrow PRIVATE XLNT_EXPORT=1)
else()
	add_library(xlntarrow STATIC ${XLNT_ARROW})
endif()

target_link_libraries(xlntarrow PRIVATE xlnt)
target_link_libraries(xlntarrow PRIVATE arrow)

target_include_directories(xlntarrow PRIVATE ${CMAKE_CURRENT_SOURCE_DIR})
target_include_directories(xlntarrow PRIVATE ${CONDA_ROOT}/include)

