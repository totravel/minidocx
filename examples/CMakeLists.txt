
link_libraries(minidocx)

add_executable(myapp main.cpp)
add_executable(pict picture.cpp)
add_executable(table table.cpp)
add_executable(style style.cpp)
add_executable(list list.cpp)

set_target_properties(
  myapp
  pict
  table
  style
  list
  PROPERTIES FOLDER examples)
