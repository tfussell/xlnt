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
    targetdir "../../bin"
    includedirs { 
       "../../include",
       "../../third-party/cxxtest",
       "../../third-party/miniz",
       "../../third-party/pugixml/src"
    }
    files { 
       "../../tests/*.hpp",
       "../../tests/runner-autogen.cpp"
    }
    links { "xlnt", "miniz" }
    prebuildcommands { "../../../third-party/cxxtest/bin/cxxtestgen --runner=ErrorPrinter -o ../../../tests/runner-autogen.cpp ../../../tests/*.hpp" }
    flags { "Unicode" }
    configuration "windows"
        defines { "WIN32" }
	links { "Shlwapi" }
    configuration "linux"
        buildoptions {
	    "-std=c++1y"
    }	
    configuration "macos"
        buildoptions {
	    "-std=c++14"
    }	

project "xlnt"
    kind "StaticLib"
    language "C++"
    targetdir "../../lib/"
    includedirs { 
       "../../include",
       "../../source",
       "../../third-party/miniz",
       "../../third-party/pugixml/src"
    }
    files {
       "../../source/**.cpp",
       "../../source/**.hpp",
       "../../include/xlnt/**.hpp",
       "../../third-party/pugixml/src/pugixml.cpp"
    }
    flags { "Unicode" }
    configuration "Debug"
        flags { "FatalWarnings" }
    configuration "windows"
        defines { 
	   "WIN32",
	   "_CRT_SECURE_NO_WARNINGS"
	}
    configuration "linux"
        buildoptions {
	    "-std=c++1y"
    }	
    configuration "macos"
        buildoptions {
	    "-std=c++14"
    }	

project "miniz"
    kind "StaticLib"
    language "C"
    targetdir "../../lib/"
    includedirs { 
       "../../third-party/miniz",
    }
    files {
       "../../third-party/miniz/miniz.c"
    }
    flags { "Unicode" }
    configuration "Debug"
        flags { "FatalWarnings" }
    configuration "windows"
        defines { 
	   "WIN32",
	   "_CRT_SECURE_NO_WARNINGS"
	}