:: gsheet->json->docx pipeline

@echo off

:: parameters
set DOCUMENT=%1

pushd .\src
.\docx-from-gsheet.py --config "../conf/config.yml" --gsheet "%DOCUMENT%"

if errorlevel 1 (
  popd
  exit /b %errorlevel%
)

popd
