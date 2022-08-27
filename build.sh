#!/bin/bash
set -e

if [ "$OSTYPE" == "msys" ]; then
  script_dir=$(dirname $(readlink -f $0) | sed 's#^/\(.\)/#\U\1:/#')
else
  script_dir=$(dirname $(readlink -f $0))
fi
sources_dir=$script_dir

if [ "$OSTYPE" == "msys" ]; then
  build_dir=$sources_dir/build-win
else
  build_dir=$sources_dir/build-linux
fi

build_type=Release # Release or Debug

minidocx_dir=$sources_dir
zip_dir=$sources_dir/../3rd_party/zip-0.2.1
pugixml_dir=$sources_dir/../3rd_party/pugixml-1.12.1

options="
  -DMINIDOCX_DIR=$minidocx_dir
  -DZIP_DIR=$zip_dir
  -DPUGIXML_DIR=$pugixml_dir"

echo -e "\e[36mBuilding...\e[0m"

if [ ! -d "$build_dir" ]; then
  mkdir $build_dir
fi

if [ "$OSTYPE" == "msys" ]; then
  cmake -S $sources_dir -B $build_dir $options
  cmake --build $build_dir --config $build_type -j4
else
  cmake -S $sources_dir -B $build_dir -DCMAKE_BUILD_TYPE=$build_type $options
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
else
  $build_dir/basic
  $build_dir/traverse
  $build_dir/breaks
  $build_dir/spacing_indent
  $build_dir/paragraph
  $build_dir/section
  $build_dir/run
fi
