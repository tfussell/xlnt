cmake_minimum_required(VERSION 3.2)
project(xlnt.third-party)

add_subdirectory(${CMAKE_CURRENT_SOURCE_DIR}/libstudxml)
add_subdirectory(${CMAKE_CURRENT_SOURCE_DIR}/cxxtest)

add_custom_target(xlnt.third-party
    DEPENDS libstudxml cxxtest-download)
