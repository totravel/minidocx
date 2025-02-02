
cmake_minimum_required(VERSION 3.28)
project(minidocx-next VERSION 1.0.0 LANGUAGES C CXX)

option(BUILD_EXAMPLES  "Build examples"               ON)
option(BUILD_TESTS     "Build tests"                  OFF)
option(WITH_STATIC_CRT "Use static C Runtime library" OFF)

set(CMAKE_C_STANDARD 11)
set(CMAKE_C_STANDARD_REQUIRED ON)
set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_DEBUG_POSTFIX d)
set(CMAKE_CONFIGURATION_TYPES Debug Release)
set(CMAKE_VS_DEBUGGER_WORKING_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}/out")

set(THIRD_PARTY_DIR "${CMAKE_CURRENT_SOURCE_DIR}/3rdparty" CACHE PATH "third-party")
set(MINIZ_DIR "${THIRD_PARTY_DIR}/miniz-3.0.2" CACHE PATH "miniz")
set(PUGIXML_DIR "${THIRD_PARTY_DIR}/pugixml-1.13" CACHE PATH "pugixml")

add_library(miniz INTERFACE)
set_target_properties(miniz PROPERTIES
  INTERFACE_INCLUDE_DIRECTORIES "${MINIZ_DIR}"
)

add_library(pugixml INTERFACE)
set_target_properties(pugixml PROPERTIES
  INTERFACE_INCLUDE_DIRECTORIES "${PUGIXML_DIR}"
  INTERFACE_SOURCES             "${PUGIXML_DIR}/pugixml.cpp"
)

include(cmake/util.cmake)

add_library(minidocx STATIC)
aux_inc_directory(include headers)
aux_src_directory(src sources)
source_group(TREE "${CMAKE_CURRENT_SOURCE_DIR}/include/minidocx" PREFIX "Header Files" FILES ${headers})
source_group(TREE "${CMAKE_CURRENT_SOURCE_DIR}/src" PREFIX "Source Files" FILES ${sources})
target_sources(minidocx PRIVATE ${headers} ${sources})
target_include_directories(minidocx PRIVATE include/minidocx)
target_include_directories(minidocx INTERFACE include)
target_compile_definitions(minidocx PRIVATE NAMESPACE=md)
target_link_libraries(minidocx PRIVATE miniz pugixml)
if(WITH_STATIC_CRT)
  set_target_properties(minidocx PROPERTIES MSVC_RUNTIME_LIBRARY MultiThreaded$<$<CONFIG:Debug>:Debug>)
endif()

if(BUILD_EXAMPLES)
  add_subdirectory(examples)
  set_directory_properties(PROPERTIES VS_STARTUP_PROJECT myapp)
endif()
