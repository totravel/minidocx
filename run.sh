#!/bin/bash
set -e

echo -e "\e[36mBuilding...\e[0m"

build_type=Release
build_dir=build

if [ "$OSTYPE" == "msys" ]; then
  cmake -S . -B $build_dir
  cmake --build $build_dir --config $build_type -j4
else
  cmake -S . -B $build_dir -DCMAKE_BUILD_TYPE=$build_type
  cmake --build $build_dir -j4
fi

echo -e "\e[36mRunning...\e[0m"

if [ "$OSTYPE" == "msys" ]; then
  $build_dir/$build_type/basic
  $build_dir/$build_type/traverse
  $build_dir/$build_type/breaks
  $build_dir/$build_type/spacing_indent
  $build_dir/$build_type/paragraph
  $build_dir/$build_type/section
  $build_dir/$build_type/run
  $build_dir/$build_type/table
  $build_dir/$build_type/table_advanced
  $build_dir/$build_type/text_frame
  $build_dir/$build_type/page_num
else
  $build_dir/basic
  $build_dir/traverse
  $build_dir/breaks
  $build_dir/spacing_indent
  $build_dir/paragraph
  $build_dir/section
  $build_dir/run
  $build_dir/table
  $build_dir/table_advanced
  $build_dir/text_frame
  $build_dir/page_num
fi
