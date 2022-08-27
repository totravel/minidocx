
cmake_minimum_required(VERSION 3.21)
project(myproj VERSION 0.1.0 LANGUAGES C CXX) # C needed by zip.c

add_library(minidocx INTERFACE)
set_target_properties(minidocx PROPERTIES
  INTERFACE_INCLUDE_DIRECTORIES "${MINIDOCX_DIR}/src"
  INTERFACE_SOURCES             "${MINIDOCX_DIR}/src/minidocx.cpp"
  INTERFACE_COMPILE_OPTIONS     "$<$<CXX_COMPILER_ID:MSVC>:/utf-8>"
)

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

target_link_libraries(minidocx INTERFACE zip pugixml)

link_libraries(minidocx)
add_executable(basic          examples/basic.cpp)
add_executable(traverse       examples/traverse.cpp)
add_executable(breaks         examples/breaks.cpp)
add_executable(spacing_indent examples/spacing_indent.cpp)
add_executable(paragraph      examples/paragraph.cpp)
add_executable(section        examples/section.cpp)
add_executable(run            examples/run.cpp)
