#!/bin/bash
set -e

if [ "$OSTYPE" == "msys" ]; then
  script_dir=$(dirname $(readlink -f $0) | sed 's#^/\(.\)/#\U\1:/#')
else
  script_dir=$(dirname $(readlink -f $0))
fi
parent_dir=$(dirname $script_dir)

if [ "$OSTYPE" == "msys" ]; then
  build_dir=$script_dir/build-win
else
  build_dir=$script_dir/build-linux
fi

build_type=Release
build_shared_libs=ON
# Windows
host=x64 # x64 or x86
arch=x64 # x64 or Win32
# Linux
build_x86=OFF

zip_dir=$parent_dir/3rd_party/zip-0.2.1
pugixml_dir=$parent_dir/3rd_party/pugixml-1.12.1
minidocx_dir=$script_dir

options="
  -DBUILD_SHARED_LIBS=$build_shared_libs
  -DZIP_DIR=$zip_dir
  -DPUGIXML_DIR=$pugixml_dir
  -DMINIDOCX_DIR=$minidocx_dir"

echo -e "\e[36mBuilding...\e[0m"

if [ ! -d "$build_dir" ]; then
  mkdir $build_dir
fi

if [ "$OSTYPE" == "msys" ]; then
  cmake -S $script_dir -B $build_dir -Thost=$host -A $arch $options
  cmake --build $build_dir --config $build_type -- -m:4
else
  cmake -S $script_dir -B $build_dir -DCMAKE_BUILD_TYPE=$build_type -DBUILD_X86=$build_x86 $options
  cmake --build $build_dir -- -j4
fi
