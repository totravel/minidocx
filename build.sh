#!/bin/bash
set -e

if [ "$OSTYPE" == "msys" ]; then
  script_dir=$(dirname $(readlink -f $0) | sed 's#^/\(.\)/#\U\1:/#')
else
  script_dir=$(dirname $(readlink -f $0))
fi
parent_dir=$(dirname $script_dir)

build_type=Release
build_shared_libs=ON
arch=x64 # or Win32

install_dir=install
zip_dir=$parent_dir/3rd_party/zip-0.2.1
pugixml_dir=$parent_dir/3rd_party/pugixml-1.12.1
minidocx_dir=$script_dir

mkdir -p build
cd build
# rm -rf *

options="
  -DBUILD_SHARED_LIBS=$build_shared_libs
  -DZIP_DIR=$zip_dir
  -DPUGIXML_DIR=$pugixml_dir
  -DMINIDOCX_DIR=$minidocx_dir"

if [ "$OSTYPE" == "msys" ]; then
  cmake .. $options -A $arch
  cmake --build . --config $build_type -- -m:4
  cmake --install . --prefix $install_dir --config $build_type
else
  cmake .. $options -DCMAKE_BUILD_TYPE=$build_type
  cmake --build . -- -j4
  cmake --install . --prefix $install_dir
fi
