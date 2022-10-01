
cmake_minimum_required(VERSION 3.21)
project(myproj VERSION 0.1.0 LANGUAGES C CXX) # C needed by zip.c

option(BUILD_EXAMPLES  "Build examples"               ON)
option(WITH_STATIC_CRT "Use static C Runtime library" OFF)

set(THIRD_PARTY_DIR "${CMAKE_CURRENT_SOURCE_DIR}/3rdparty" CACHE PATH "3rdparty")
set(ZIP_DIR         "${THIRD_PARTY_DIR}/zip-0.2.1"         CACHE PATH "zip")
set(PUGIXML_DIR     "${THIRD_PARTY_DIR}/pugixml-1.12.1"    CACHE PATH "pugixml")

set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_DEBUG_POSTFIX d)

add_library(zip INTERFACE)
set_target_properties(zip PROPERTIES
  INTERFACE_INCLUDE_DIRECTORIES "${ZIP_DIR}/src"
  INTERFACE_SOURCES             "${ZIP_DIR}/src/zip.c"
)

add_library(pugixml INTERFACE)
set_target_properties(pugixml PROPERTIES
  INTERFACE_INCLUDE_DIRECTORIES "${PUGIXML_DIR}/src"
  INTERFACE_SOURCES             "${PUGIXML_DIR}/src/pugixml.cpp"
)

add_library(minidocx STATIC)
target_sources(minidocx PRIVATE "${CMAKE_CURRENT_SOURCE_DIR}/src/minidocx.cpp")
target_include_directories(minidocx PUBLIC "${CMAKE_CURRENT_SOURCE_DIR}/src")
target_include_directories(minidocx PUBLIC "${PUGIXML_DIR}/src")
target_link_libraries(minidocx PRIVATE zip pugixml)
if(WITH_STATIC_CRT)
  set_target_properties(minidocx 
    PROPERTIES 
      MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")
endif()

if(BUILD_EXAMPLES)
  link_libraries(minidocx)
  add_executable(basic          examples/basic.cpp)
  add_executable(traverse       examples/traverse.cpp)
  add_executable(breaks         examples/breaks.cpp)
  add_executable(spacing_indent examples/spacing_indent.cpp)
  add_executable(paragraph      examples/paragraph.cpp)
  add_executable(section        examples/section.cpp)
  add_executable(run            examples/run.cpp)
endif()

install(TARGETS minidocx)
install(FILES "${CMAKE_CURRENT_SOURCE_DIR}/src/minidocx.hpp" TYPE INCLUDE)
install(FILES "${PUGIXML_DIR}/src/pugixml.hpp" TYPE INCLUDE)
