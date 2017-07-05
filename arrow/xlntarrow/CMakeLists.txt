cmake_minimum_required(VERSION 3.2)
project(xlntarrow)

set(CMAKE_CXX_STANDARD 14)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(XLNT_ARROW
    ${CMAKE_CURRENT_SOURCE_DIR}/xlntarrow.hpp
    ${CMAKE_CURRENT_SOURCE_DIR}/xlntarrow.cpp)

add_library(xlntarrow SHARED ${XLNT_ARROW})
target_compile_definitions(xlntarrow PRIVATE XLNT_EXPORT=1)
target_link_libraries(xlntarrow PRIVATE xlnt)
link_directories(${ARROW_LIBRARY_PATH})
target_link_libraries(xlntarrow PRIVATE arrow)
target_include_directories(xlntarrow PRIVATE ${CMAKE_CURRENT_SOURCE_DIR})
target_include_directories(xlntarrow PRIVATE ${ARROW_INCLUDE_PATH})
target_include_directories(xlntarrow PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/../../include)
