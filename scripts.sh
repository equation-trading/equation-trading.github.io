#!/bin/bash

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"

# run: bundle exec jekyll serve -H 0.0.0.0

cd $SCRIPT_DIR

bundle exec jekyll build --watch
