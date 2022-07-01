#!/bin/bash
set -e

./build.sh

./build/install/bin/basic
./build/install/bin/traverse
./build/install/bin/breaks
./build/install/bin/spacing_indent
./build/install/bin/paragraph
./build/install/bin/section
./build/install/bin/paragraph_type
./build/install/bin/run
