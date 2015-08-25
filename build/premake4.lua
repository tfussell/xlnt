solution "xlnt"
    configurations { "Debug", "Release" }
    platforms { "x64" }
    location ("./" .. _ACTION)
    configuration "Debug"
        flags { "Symbols" }

project "xlnt.test"
    kind "ConsoleApp"
    language "C++"
    targetname "xlnt.test"
    targetdir "../bin"
    includedirs { 
       "../include",
       "../third-party/pugixml/src",
       "../third-party/miniz",
       "/usr/local/Cellar/cxxtest/4.3"
    }
    files { 
       "../tests/*.hpp",
       "../tests/runner-autogen.cpp"
    }
    links { "xlnt" }
    prebuildcommands { "/usr/local/Cellar/cxxtest/4.3/bin/cxxtestgen --runner=ErrorPrinter -o ../../tests/runner-autogen.cpp ../../tests/*.hpp" }
    flags { "Unicode" }
    configuration "windows"
        defines { "WIN32" }
	links { "Shlwapi" }

project "xlnt"
    kind "StaticLib"
    language "C++"
    targetdir "../lib/"
    includedirs { 
       "../include",
       "../third-party/pugixml/src",
       "../third-party/miniz"
    }
    files {
       "../source/**.cpp",
       "../source/**.hpp",
       "../include/xlnt/**.hpp",
       "../third-party/pugixml/src/pugixml.cpp",
       "../third-party/miniz/miniz.c"
    }
    flags { "Unicode" }
    configuration "Debug"
        flags { "FatalWarnings" }
    configuration "windows"
        defines { 
	   "WIN32",
	   "_CRT_SECURE_NO_WARNINGS"
	}
