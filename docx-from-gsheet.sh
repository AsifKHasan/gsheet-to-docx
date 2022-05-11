#!/usr/bin/env bash

# parameters
DOCUMENT=$1

set echo off

pushd ./src
./docx-from-gsheet.py --config "../conf/config.yml" --gsheet ${DOCUMENT}

if [ ${?} -ne 0 ]; then
  popd && exit 1
else
  popd
fi
