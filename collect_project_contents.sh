#!/bin/bash

find . -type f \
    -not -path '*/\.*' \
    -not -name '.clasp.json' \
    -not -name '.claspignore' \
    -not -name 'Dockerfile' \
    -not -name '*.txt' \
    -not -name '*.md' \
    -not -name '*.yaml' \
    -not -name '.env' \
    -not -name '.sh' \
    \( -name '*.gs' -o -name '*.html' -o -name 'appsscript.json' \) | 
while read file; do 
    echo "# File: $file"
    echo
    cat "$file"
    echo
    echo
done > project_contents.txt