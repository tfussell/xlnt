project(xlnt)

set(PROJECT_VENDOR "Thomas Fussell")
set(PROJECT_CONTACT "thomas.fussellgmail.com")
set(PROJECT_URL "https://github.com/tfussell/xlnt")
set(PROJECT_DESCRIPTION "user-friendly xlsx library for C++14")
include(VERSION.cmake)

if(NOT CMAKE_INSTALL_PREFIX)
  if(MSVC)
    set(CMAKE_INSTALL_PREFIX /c/Program Files/xlnt)
  else()
    set(CMAKE_INSTALL_PREFIX /usr/local)
  endif()
endif()

set(INC_DEST_DIR ${CMAKE_INSTALL_PREFIX}/include)

if(NOT LIB_DEST_DIR)
  set(LIB_DEST_DIR ${CMAKE_INSTALL_PREFIX}/lib)
endif()

if(NOT BIN_DEST_DIR)
  set(BIN_DEST_DIR ${CMAKE_INSTALL_PREFIX}/bin)
endif()

include_directories(../include)
include_directories(../include/xlnt)
include_directories(../source)
include_directories(../third-party/miniz)
include_directories(../third-party/pugixml/src)
include_directories(../third-party/utfcpp/source)

FILE(GLOB ROOT_HEADERS ../include/xlnt/*.hpp)
FILE(GLOB CELL_HEADERS ../include/xlnt/cell/*.hpp)
FILE(GLOB CHARTS_HEADERS ../include/xlnt/charts/*.hpp)
FILE(GLOB CHARTSHEET_HEADERS ../include/xlnt/chartsheet/*.hpp)
FILE(GLOB DRAWING_HEADERS ../include/xlnt/drawing/*.hpp)
FILE(GLOB FORMULA_HEADERS ../include/xlnt/formula/*.hpp)
FILE(GLOB PACKAGING_HEADERS ../include/xlnt/packaging/*.hpp)
FILE(GLOB SERIALIZATION_HEADERS ../include/xlnt/serialization/*.hpp)
FILE(GLOB STYLES_HEADERS ../include/xlnt/styles/*.hpp)
FILE(GLOB UTILS_HEADERS ../include/xlnt/utils/*.hpp)
FILE(GLOB WORKBOOK_HEADERS ../include/xlnt/workbook/*.hpp)
FILE(GLOB WORKSHEET_HEADERS ../include/xlnt/worksheet/*.hpp)
FILE(GLOB DETAIL_HEADERS ../source/detail/*.hpp)

SET(HEADERS ${ROOT_HEADERS} ${CELL_HEADERS} ${CHARTS_HEADERS} ${CHARTSHEET_HEADERS} ${DRAWING_HEADERS} ${FORMULA_HEADERS} ${PACKAGING_HEADERS} ${SERIALIZATION_HEADERS} ${STYLES_HEADERS} ${UTILS_HEADERS} ${WORKBOOK_HEADERS} ${WORKSHEET_HEADERS} ${DETAIL_HEADERS})

FILE(GLOB CELL_SOURCES ../source/cell/*.cpp)
FILE(GLOB CHARTS_SOURCES ../source/charts/*.cpp)
FILE(GLOB CHARTSHEET_SOURCES ../source/chartsheet/*.cpp)
FILE(GLOB DRAWING_SOURCES ../source/drawing/*.cpp)
FILE(GLOB FORMULA_SOURCES ../source/formula/*.cpp)
FILE(GLOB PACKAGING_SOURCES ../source/packaging/*.cpp)
FILE(GLOB SERIALIZATION_SOURCES ../source/serialization/*.cpp)
FILE(GLOB STYLES_SOURCES ../source/styles/*.cpp)
FILE(GLOB UTILS_SOURCES ../source/utils/*.cpp)
FILE(GLOB WORKBOOK_SOURCES ../source/workbook/*.cpp)
FILE(GLOB WORKSHEET_SOURCES ../source/worksheet/*.cpp)
FILE(GLOB DETAIL_SOURCES ../source/detail/*.cpp)

SET(SOURCES ${CELL_SOURCES} ${CHARTS_SOURCES} ${CHARTSHEET_SOURCES} ${DRAWING_SOURCES} ${FORMULA_SOURCES} ${PACKAGING_SOURCES} ${SERIALIZATION_SOURCES} ${STYLES_SOURCES} ${UTILS_SOURCES} ${WORKBOOK_SOURCES} ${WORKSHEET_SOURCES} ${DETAIL_SOURCES})

SET(MINIZ ../third-party/miniz/miniz.c ../third-party/miniz/miniz.h)

SET(PUGIXML ../third-party/pugixml/src/pugixml.hpp ../third-party/pugixml/src/pugixml.cpp ../third-party/pugixml/src/pugiconfig.hpp)

if(SHARED)
    add_library(xlnt.shared SHARED ${HEADERS} ${SOURCES} ${MINIZ} ${PUGIXML})
    add_definitions(-DXLNT_SHARED)
    if(MSVC)
        target_compile_definitions(xlnt.shared PRIVATE XLNT_EXPORT=1)
        set_target_properties(xlnt.shared PROPERTIES COMPILE_FLAGS "/wd\"4251\" /wd\"4275\"")
    endif()
    install(TARGETS xlnt.shared
        LIBRARY DESTINATION ${LIB_DEST_DIR}
        ARCHIVE DESTINATION ${LIB_DEST_DIR}
        RUNTIME DESTINATION ${BIN_DEST_DIR}
    )
    SET_TARGET_PROPERTIES(
        xlnt.shared
	PROPERTIES
	OUTPUT_NAME xlnt
	VERSION ${PROJECT_VERSION_FULL}
	SOVERSION ${PROJECT_VERSION}
	INSTALL_NAME_DIR "${LIB_DEST_DIR}"
    )

    if(FRAMEWORK)
        add_custom_command(
            TARGET xlnt.shared
            POST_BUILD
            COMMAND mkdir -p "${CMAKE_CURRENT_BINARY_DIR}/${PROJECT_NAME}.framework/Versions/${PROJECT_VERSION_FULL}/Headers"
            COMMAND cp -R ../include/xlnt/* "${CMAKE_CURRENT_BINARY_DIR}/${PROJECT_NAME}.framework/Versions/${PROJECT_VERSION_FULL}/Headers"
            COMMAND cp "../lib/lib${PROJECT_NAME}.${PROJECT_VERSION_FULL}.dylib" "${CMAKE_CURRENT_BINARY_DIR}/${PROJECT_NAME}.framework/Versions/${PROJECT_VERSION_FULL}/xlnt"
            COMMAND cd "${CMAKE_CURRENT_BINARY_DIR}/${PROJECT_NAME}.framework/Versions" && ln -s "${PROJECT_VERSION_FULL}" Current
            COMMAND cd "${CMAKE_CURRENT_BINARY_DIR}/${PROJECT_NAME}.framework" && ln -s Versions/Current/* ./
        )
    endif()
endif()

if(STATIC)
    add_library(xlnt.static STATIC ${HEADERS} ${SOURCES} ${MINIZ} ${PUGIXML})
    install(TARGETS xlnt.static
        LIBRARY DESTINATION ${LIB_DEST_DIR}
        ARCHIVE DESTINATION ${LIB_DEST_DIR}
        RUNTIME DESTINATION ${BIN_DEST_DIR}
    )
    SET_TARGET_PROPERTIES(
        xlnt.static
	PROPERTIES
	OUTPUT_NAME xlnt
    )
endif()

source_group(xlnt FILES ${ROOT_HEADERS})
source_group(detail FILES ${DETAIL_HEADERS} ${DETAIL_SOURCES})
source_group(cell FILES ${CELL_HEADERS} ${CELL_SOURCES})
source_group(charts FILES ${CHARTS_HEADERS} ${CHARTS_SOURCES})
source_group(chartsheet FILES ${CHARTSHEET_HEADERS} ${CHARTSHEET_SOURCES})
source_group(drawing FILES ${DRAWING_HEADERS} ${DRAWING_SOURCES})
source_group(formula FILES ${FORMULA_HEADERS} ${FORMULA_SOURCES})
source_group(packaging FILES ${PACKAGING_HEADERS} ${PACKAGING_SOURCES})
source_group(serialization FILES ${SERIALIZATION_HEADERS} ${SERIALIZATION_SOURCES})
source_group(styles FILES ${STYLES_HEADERS} ${STYLES_SOURCES})
source_group(utils FILES ${UTILS_HEADERS} ${UTILS_SOURCES})
source_group(workbook FILES ${WORKBOOK_HEADERS} ${WORKBOOK_SOURCES})
source_group(worksheet FILES ${WORKSHEET_HEADERS} ${WORKSHEET_SOURCES})
source_group(third-party\\miniz FILES ${MINIZ})
source_group(third-party\\pugixml FILES ${PUGIXML})

SET(CMAKE_INSTALL_RPATH_USE_LINK_PATH TRUE)

SET(PKG_CONFIG_LIBDIR ${LIB_DEST_DIR})
SET(PKG_CONFIG_INCLUDEDIR ${INC_DEST_DIR})
SET(PKG_CONFIG_LIBS "-L\${libdir} -lxlnt")
SET(PKG_CONFIG_CFLAGS "-I\${includedir}")

CONFIGURE_FILE(
  "${CMAKE_CURRENT_SOURCE_DIR}/pkg-config.pc.cmake"
  "${CMAKE_CURRENT_BINARY_DIR}/${PROJECT_NAME}.pc"
)

configure_file(
    "${CMAKE_CURRENT_SOURCE_DIR}/cmake_uninstall.cmake.in"
    "${CMAKE_CURRENT_BINARY_DIR}/cmake_uninstall.cmake"
    IMMEDIATE @ONLY)

install(DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}/../include/xlnt 
        DESTINATION include
        PATTERN ".DS_Store" EXCLUDE
)

install(FILES "${CMAKE_BINARY_DIR}/${PROJECT_NAME}.pc"
        DESTINATION ${LIB_DEST_DIR}/pkgconfig
)

add_custom_target(uninstall
    COMMAND ${CMAKE_COMMAND} -P ${CMAKE_CURRENT_BINARY_DIR}/cmake_uninstall.cmake)