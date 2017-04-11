cmake_minimum_required(VERSION 3.2)
project(xlnt.benchmarks)

set(CMAKE_CXX_STANDARD 14)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

if(NOT COMBINED_PROJECT)
    add_subdirectory(${CMAKE_CURRENT_SOURCE_DIR}/../source ${CMAKE_CURRENT_BINARY_DIR}/source)
endif()

include_directories(${LIBRARY_INCLUDE_DIR})

file(GLOB SAMPLE_SOURCES ${CMAKE_CURRENT_SOURCE_DIR}/*.cpp)

foreach(SAMPLE_SOURCE IN ITEMS ${SAMPLE_SOURCES})
    get_filename_component(SAMPLE_NAME ${SAMPLE_SOURCE} NAME_WE)
    set(SAMPLE_EXECUTABLE benchmark-${SAMPLE_NAME})
    add_executable(${SAMPLE_EXECUTABLE} ${SAMPLE_SOURCE})
    target_link_libraries(${SAMPLE_EXECUTABLE} PRIVATE xlnt)
    if(NOT STATIC)
        add_custom_command(TARGET ${SAMPLE_EXECUTABLE} POST_BUILD 
            COMMAND ${CMAKE_COMMAND} -E copy_if_different
            $<TARGET_FILE:xlnt>
            $<TARGET_FILE_DIR:${SAMPLE_EXECUTABLE}>)
    endif()    
endforeach()
