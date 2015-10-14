cmake_minimum_required(VERSION 2.8.9)
project(xlnt)

include_directories(../../../include)
include_directories(../../../third-party/miniz)
include_directories(../../../third-party/pugixml/src)

FILE(GLOB SOURCES ../../../source/*.cpp)
FILE(GLOB DETAIL_SOURCES ../../../source/detail/*.cpp)
add_library(xlnt STATIC ${SOURCES} ${DETAIL_SOURCES} ../../../third-party/pugixml/src/pugixml.cpp ../../../third-party/miniz/miniz.c)

