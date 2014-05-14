solution "xlnt"
    configurations { "Debug", "Release" }
    platforms { "x64" }
    location ("./" .. _ACTION)
    configuration "not windows"
        buildoptions { 
            "-std=c++11",
            "-Wno-unknown-pragmas"
        }
    configuration "Debug"
        flags { "Symbols" }
	optimize "Off"
    configuration "Release"
        optimize "Full"

project "xlnt.test"
    kind "ConsoleApp"
    language "C++"
    targetname "xlnt.test"
    includedirs { 
       "$(cxxtest_prefix)",
       "../include",
       "../third-party/pugixml/src",
       "../third-party/zlib",
       "../third-party/zlib/contrib/minizip"
    }
    defines { "WIN32" }
    files { 
       "../source/tests/**.h",
       "../source/tests/runner-autogen.cpp"
    }
    links { "xlnt" }
    prebuildcommands { "cxxtestgen --runner=ErrorPrinter -o ../../source/tests/runner-autogen.cpp ../../source/tests/*.h" }
    flags { 
       "Unicode",
       "NoEditAndContinue",
       "NoManifest",
       "NoPCH"
    }
    configuration "Debug"
	targetdir "../bin"
    configuration "Release"
        flags { "LinkTimeOptimization" }
	targetdir "../bin"
    configuration "windows"
        defines { "WIN32" }
	links { "Shlwapi" }

project "xlnt"
    kind "StaticLib"
    language "C++"
    warnings "Extra"
    targetdir "../lib/"
    includedirs { 
       "../include/xlnt",
       "../third-party/pugixml/src",
       "../source/",
       "../third-party/zlib/",
       "../third-party/zlib/contrib/minizip"
    }
    files {
       "../source/**.cpp",
       "../source/**.h",
       "../include/**.h",
       "../third-party/pugixml/src/pugixml.cpp",
       "../third-party/zlib/*.c",
       "../third-party/zlib/contrib/minizip/*.c"
    }
    excludes {
       "../source/tests/**.cpp",
       "../source/tests/**.h",
       "../third-party/zlib/contrib/minizip/miniunz.c",
       "../third-party/zlib/contrib/minizip/minizip.c"
    }
    flags { 
       "Unicode",
       "NoEditAndContinue",
       "NoManifest",
       "NoPCH"
    }
    configuration "Debug"
        flags { "FatalWarnings" }
    configuration "windows"
        defines { "WIN32" }
    configuration "../third-party/zlib/*.c"
        warnings "Off"
