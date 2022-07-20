#!/bin/bash
set -e

. ./build.sh

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
