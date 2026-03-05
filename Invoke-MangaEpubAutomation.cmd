@echo off
setlocal
REM Default behavior now includes preflight plan summary and confirmation gate.
REM Use -AutoConfirm for non-interactive/scheduled runs.
REM Use -InitDepsConfig to generate manga_epub_automation.deps.json template.
REM Stage switches: -SkipUpscale / -SkipEpubPackaging / -SkipMergedEpub
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Invoke-MangaEpubAutomation.ps1" %*
set EXITCODE=%ERRORLEVEL%
exit /b %EXITCODE%

