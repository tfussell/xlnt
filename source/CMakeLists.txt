if(STATIC)
	set(PROJECT_NAME xlnt.static)
else()
	set(PROJECT_NAME xlnt.shared)
endif()

project(${PROJECT_NAME})

set(PROJECT_VENDOR "Thomas Fussell")
set(PROJECT_CONTACT "thomas.fussell@gmail.com")
set(PROJECT_URL "https://github.com/tfussell/xlnt")
set(PROJECT_DESCRIPTION "cross-platform user-friendly xlsx library for C++14")

if(CMAKE_INSTALL_PREFIX_INITIALIZED_TO_DEFAULT)
    set(CMAKE_INSTALL_PREFIX "${CMAKE_BINARY_DIR}/installed" CACHE PATH "default install path" FORCE)
endif()

if(NOT INC_DEST_DIR)
    set(INC_DEST_DIR ${CMAKE_INSTALL_PREFIX}/include)
endif()

if(NOT LIB_DEST_DIR)
    set(LIB_DEST_DIR ${CMAKE_INSTALL_PREFIX}/lib)
endif()

if(NOT BIN_DEST_DIR)
    set(BIN_DEST_DIR ${CMAKE_INSTALL_PREFIX}/bin)
endif()

include_directories(${XLNT_INCLUDE_DIR})
include_directories(${XLNT_INCLUDE_DIR}/xlnt)
include_directories(${XLNT_SOURCE_DIR})
include_directories(${XLNT_SOURCE_DIR}/detail)
include_directories(${THIRD_PARTY_DIR}/miniz)
include_directories(${THIRD_PARTY_DIR}/libstudxml)
include_directories(${THIRD_PARTY_DIR}/utfcpp/source)
include_directories(${THIRD_PARTY_DIR}/pole)
include_directories(${THIRD_PARTY_DIR}/botan)

file(GLOB ROOT_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/*.hpp)
file(GLOB CELL_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/cell/*.hpp)
file(GLOB CHARTS_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/charts/*.hpp)
file(GLOB CHARTSHEET_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/chartsheet/*.hpp)
file(GLOB DRAWING_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/drawing/*.hpp)
file(GLOB FORMULA_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/formula/*.hpp)
file(GLOB PACKAGING_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/packaging/*.hpp)
file(GLOB STYLES_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/styles/*.hpp)
file(GLOB UTILS_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/utils/*.hpp)
file(GLOB WORKBOOK_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/workbook/*.hpp)
file(GLOB WORKSHEET_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/worksheet/*.hpp)
file(GLOB DETAIL_HEADERS ${XLNT_INCLUDE_DIR}/detail/*.hpp)

set(HEADERS ${ROOT_HEADERS} ${CELL_HEADERS} ${CHARTS_HEADERS} 
    ${CHARTSHEET_HEADERS} ${DRAWING_HEADERS} ${FORMULA_HEADERS}
	${PACKAGING_HEADERS} ${STYLES_HEADERS} ${UTILS_HEADERS} 
	${WORKBOOK_HEADERS} ${WORKSHEET_HEADERS} ${DETAIL_HEADERS})

file(GLOB CELL_SOURCES ${XLNT_SOURCE_DIR}/cell/*.cpp)
file(GLOB CHARTS_SOURCES ${XLNT_SOURCE_DIR}/charts/*.cpp)
file(GLOB CHARTSHEET_SOURCES ${XLNT_SOURCE_DIR}/chartsheet/*.cpp)
file(GLOB DRAWING_SOURCES ${XLNT_SOURCE_DIR}/drawing/*.cpp)
file(GLOB FORMULA_SOURCES ${XLNT_SOURCE_DIR}/formula/*.cpp)
file(GLOB PACKAGING_SOURCES ${XLNT_SOURCE_DIR}/packaging/*.cpp)
file(GLOB STYLES_SOURCES ${XLNT_SOURCE_DIR}/styles/*.cpp)
file(GLOB UTILS_SOURCES ${XLNT_SOURCE_DIR}/utils/*.cpp)
file(GLOB WORKBOOK_SOURCES ${XLNT_SOURCE_DIR}/workbook/*.cpp)
file(GLOB WORKSHEET_SOURCES ${XLNT_SOURCE_DIR}/worksheet/*.cpp)
file(GLOB DETAIL_SOURCES ${XLNT_SOURCE_DIR}/detail/*.cpp)

set(SOURCES ${CELL_SOURCES} ${CHARTS_SOURCES} ${CHARTSHEET_SOURCES} 
    ${DRAWING_SOURCES} ${FORMULA_SOURCES} ${PACKAGING_SOURCES} 
	${STYLES_SOURCES} ${UTILS_SOURCES} ${WORKBOOK_SOURCES} 
	${WORKSHEET_SOURCES} ${DETAIL_SOURCES})

set(MINIZ ../third-party/miniz/miniz.c ../third-party/miniz/miniz.h)
set(LIBSTUDXML ../third-party/libstudxml/xml/parser.cxx ../third-party/libstudxml/xml/qname.cxx 
    ../third-party/libstudxml/xml/serializer.cxx ../third-party/libstudxml/xml/value-traits.cxx 
	../third-party/libstudxml/xml/details/expat/xmlparse.c ../third-party/libstudxml/xml/details/expat/xmlrole.c 
	../third-party/libstudxml/xml/details/expat/xmltok_impl.c ../third-party/libstudxml/xml/details/expat/xmltok_ns.c 
	../third-party/libstudxml/xml/details/expat/xmltok.c ../third-party/libstudxml/xml/details/genx/char-props.c 
	../third-party/libstudxml/xml/details/genx/genx.c)
set(POLE ${THIRD_PARTY_DIR}/pole/pole.cpp)
set(BOTAN ${THIRD_PARTY_DIR}/botan/botan_all.cpp)

if(MSVC)
    set_source_files_properties(${BOTAN} PROPERTIES COMPILE_FLAGS "/wd\"4244\"")
endif()

if (CMAKE_SIZEOF_VOID_P EQUAL 8)
    set(CPU "x64")
else()
    set(CPU "x86")
endif()

add_custom_command(OUTPUT ${BOTAN}
    COMMAND python configure.py --minimized-build --enable-modules=sha1,aes,filters,codec_filt,cbc,ecb,sha2_32,sha2_64 --disable-shared --amalgamation --cpu=${CPU}
    WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}/third-party/botan
    COMMENT "Generating botan amalgamation ${BOTAN}")

if(SHARED)
    add_library(xlnt.shared SHARED ${HEADERS} ${SOURCES} ${MINIZ} ${LIBSTUDXML} ${POLE} ${BOTAN})
    target_compile_definitions(xlnt.shared PRIVATE XLNT_SHARED=1 LIBSTUDXML_STATIC_LIB=1)

    if(MSVC)
        target_compile_definitions(xlnt.shared PRIVATE XLNT_EXPORT=1 _CRT_SECURE_NO_WARNINGS=1 NOMINMAX)
        set_target_properties(xlnt.shared PROPERTIES COMPILE_FLAGS "/wd\"4251\" /wd\"4275\" /MP")
    endif()

    install(TARGETS xlnt.shared
        LIBRARY DESTINATION ${LIB_DEST_DIR}
        ARCHIVE DESTINATION ${LIB_DEST_DIR}
        RUNTIME DESTINATION ${BIN_DEST_DIR})
    set_target_properties(
        xlnt.shared
        PROPERTIES
        OUTPUT_NAME xlnt
        VERSION ${PROJECT_VERSION_FULL}
        SOVERSION ${PROJECT_VERSION}
        INSTALL_NAME_DIR "${LIB_DEST_DIR}"
        RUNTIME_OUTPUT_DIRECTORY_DEBUG "${CMAKE_BINARY_DIR}/bin"
        RUNTIME_OUTPUT_DIRECTORY_MINSIZEREL "${CMAKE_BINARY_DIR}/bin"
        RUNTIME_OUTPUT_DIRECTORY_RELEASE "${CMAKE_BINARY_DIR}/bin"
        RUNTIME_OUTPUT_DIRECTORY_RELWITHDEBINFO "${CMAKE_BINARY_DIR}/bin"
        ARCHIVE_OUTPUT_DIRECTORY_DEBUG "${CMAKE_BINARY_DIR}/lib"
        ARCHIVE_OUTPUT_DIRECTORY_MINSIZEREL "${CMAKE_BINARY_DIR}/lib"
        ARCHIVE_OUTPUT_DIRECTORY_RELEASE "${CMAKE_BINARY_DIR}/lib"
        ARCHIVE_OUTPUT_DIRECTORY_RELWITHDEBINFO "${CMAKE_BINARY_DIR}/lib"
        LIBRARY_OUTPUT_DIRECTORY_DEBUG "${CMAKE_BINARY_DIR}/lib"
        LIBRARY_OUTPUT_DIRECTORY_MINSIZEREL "${CMAKE_BINARY_DIR}/lib"
        LIBRARY_OUTPUT_DIRECTORY_RELEASE "${CMAKE_BINARY_DIR}/lib"
        LIBRARY_OUTPUT_DIRECTORY_RELWITHDEBINFO "${CMAKE_BINARY_DIR}/lib")

    if(FRAMEWORK)
        add_custom_command(
            TARGET xlnt.shared
            POST_BUILD
            COMMAND mkdir -p "${CMAKE_CURRENT_BINARY_DIR}/${PROJECT_NAME}.framework/Versions/${PROJECT_VERSION_FULL}/Headers"
            COMMAND cp -R ../include/xlnt/* "${CMAKE_CURRENT_BINARY_DIR}/${PROJECT_NAME}.framework/Versions/${PROJECT_VERSION_FULL}/Headers"
            COMMAND cp "lib/lib${PROJECT_NAME}.${PROJECT_VERSION_FULL}.dylib" "${CMAKE_CURRENT_BINARY_DIR}/${PROJECT_NAME}.framework/Versions/${PROJECT_VERSION_FULL}/xlnt"
            COMMAND cd "${CMAKE_CURRENT_BINARY_DIR}/${PROJECT_NAME}.framework/Versions" && ln -s "${PROJECT_VERSION_FULL}" Current
            COMMAND cd "${CMAKE_CURRENT_BINARY_DIR}/${PROJECT_NAME}.framework" && ln -s Versions/Current/* ./
        )
    endif()
endif()

if(STATIC)
    add_library(xlnt.static STATIC ${HEADERS} ${SOURCES} ${MINIZ} ${LIBSTUDXML} ${POLE} ${BOTAN})
    target_compile_definitions(xlnt.static PUBLIC XLNT_STATIC=1)
    target_compile_definitions(xlnt.static PRIVATE LIBSTUDXML_STATIC_LIB=1)

    if(MSVC)
        target_compile_definitions(xlnt.static PRIVATE _CRT_SECURE_NO_WARNINGS=1 NOMINMAX)
		set_target_properties(xlnt.static PROPERTIES COMPILE_FLAGS "/MP")
    endif()

    install(TARGETS xlnt.static
        LIBRARY DESTINATION ${LIB_DEST_DIR}
        ARCHIVE DESTINATION ${LIB_DEST_DIR}
        RUNTIME DESTINATION ${BIN_DEST_DIR}
    )
    set_target_properties(xlnt.static
        PROPERTIES
        OUTPUT_NAME xlnt
        RUNTIME_OUTPUT_DIRECTORY_DEBUG "${CMAKE_BINARY_DIR}/bin"
        RUNTIME_OUTPUT_DIRECTORY_MINSIZEREL "${CMAKE_BINARY_DIR}/bin"
        RUNTIME_OUTPUT_DIRECTORY_RELEASE "${CMAKE_BINARY_DIR}/bin"
        RUNTIME_OUTPUT_DIRECTORY_RELWITHDEBINFO "${CMAKE_BINARY_DIR}/bin"
        ARCHIVE_OUTPUT_DIRECTORY_DEBUG "${CMAKE_BINARY_DIR}/lib"
        ARCHIVE_OUTPUT_DIRECTORY_MINSIZEREL "${CMAKE_BINARY_DIR}/lib"
        ARCHIVE_OUTPUT_DIRECTORY_RELEASE "${CMAKE_BINARY_DIR}/lib"
        ARCHIVE_OUTPUT_DIRECTORY_RELWITHDEBINFO "${CMAKE_BINARY_DIR}/lib"
        LIBRARY_OUTPUT_DIRECTORY_DEBUG "${CMAKE_BINARY_DIR}/lib"
        LIBRARY_OUTPUT_DIRECTORY_MINSIZEREL "${CMAKE_BINARY_DIR}/lib"
        LIBRARY_OUTPUT_DIRECTORY_RELEASE "${CMAKE_BINARY_DIR}/lib"
        LIBRARY_OUTPUT_DIRECTORY_RELWITHDEBINFO "${CMAKE_BINARY_DIR}/lib")
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
source_group(third-party\\libstudxml FILES ${LIBSTUDXML})

set(CMAKE_INSTALL_RPATH_USE_LINK_PATH TRUE)

set(PKG_CONFIG_LIBDIR ${LIB_DEST_DIR})
set(PKG_CONFIG_INCLUDEDIR ${INC_DEST_DIR})
set(PKG_CONFIG_LIBS "-L\${libdir} -lxlnt")
set(PKG_CONFIG_CFLAGS "-I\${includedir}")

configure_file(
    "${CMAKE_SOURCE_DIR}/cmake/pkg-config.pc.cmake"
    "${CMAKE_CURRENT_BINARY_DIR}/${PROJECT_NAME}.pc")

configure_file(
    "${CMAKE_SOURCE_DIR}/cmake/cmake_uninstall.cmake.in"
    "${CMAKE_CURRENT_BINARY_DIR}/cmake_uninstall.cmake"
    IMMEDIATE @ONLY)

install(
    DIRECTORY ${XLNT_INCLUDE_DIR}/xlnt
    DESTINATION include
    PATTERN ".DS_Store" EXCLUDE)

install(
    FILES "${CMAKE_CURRENT_BINARY_DIR}/${PROJECT_NAME}.pc"
    DESTINATION ${LIB_DEST_DIR}/pkgconfig)

add_custom_target(
    uninstall
    COMMAND ${CMAKE_COMMAND} -P ${CMAKE_CURRENT_BINARY_DIR}/cmake_uninstall.cmake)
