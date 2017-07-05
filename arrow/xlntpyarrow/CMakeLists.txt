cmake_minimum_required(VERSION 3.2)
project(xlntpyarrow)

configure_file("${CMAKE_CURRENT_SOURCE_DIR}/setup.py.cmake"
    "${CMAKE_CURRENT_BINARY_DIR}/setup.py")



add_custom_command(xlntpyarrow
  COMMAND python ${CMAKE_CURRENT_BINARY_DIR}/setup.py build_ext)
