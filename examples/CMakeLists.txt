
add_executable(myapp main.cpp)
target_link_libraries(myapp PRIVATE minidocx)

add_executable(pict picture.cpp)
target_link_libraries(pict PRIVATE minidocx)

add_executable(table table.cpp)
target_link_libraries(table PRIVATE minidocx)

add_executable(list list.cpp)
target_link_libraries(list PRIVATE minidocx)

add_executable(style style.cpp)
target_link_libraries(style PRIVATE minidocx)
