
# Collects the names of all the header files in the specified directory 
# and stores the list in the <variable> provided. 
macro(aux_inc_directory dir variable)
  file(GLOB_RECURSE ${variable} LIST_DIRECTORIES false "${dir}/*.h" "${dir}/*.hpp")
endmacro()

# Collects the names of all the source files in the specified directory 
# and stores the list in the <variable> provided. 
macro(aux_src_directory dir variable)
  file(GLOB_RECURSE ${variable} LIST_DIRECTORIES false "${dir}/*.c" "${dir}/*.cpp" "${dir}/*.cc" "${dir}/*.cxx")
endmacro()
