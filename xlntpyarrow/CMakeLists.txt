cmake_minimum_required(VERSION 3.2)
project(xlntpyarrow)

add_subdirectory(../third-party/pybind11 pybind11)
set(CMAKE_MODULE_PATH ${CMAKE_MODULE_PATH} "${CMAKE_CURRENT_SOURCE_DIR}/../cmake")

find_package(Arrow)

if(NOT ARROW_FOUND)
	message(FATAL_ERROR "Arrow not found.")
endif()

pybind11_add_module(xlntpyarrow xlntpyarrow.cpp)

target_include_directories(xlntpyarrow
	PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}
	PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/../source
	PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/../third-party/pybind11/include
	PRIVATE ${ARROW_INCLUDE_DIR})
target_link_libraries(xlntpyarrow
	PRIVATE xlnt
	PRIVATE ${ARROW_SHARED_IMP_LIB}
	PRIVATE ${ARROW_PYTHON_SHARED_IMP_LIB})
