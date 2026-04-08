@echo off
setlocal

set VCVARS="C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\VC\Auxiliary\Build\vcvarsall.bat"
set SRC_DIR=%~dp0

echo === Build started ===

REM ---- 64bit ----
echo [64bit] Setting up environment...
call %VCVARS% x64 > nul 2>&1
echo [64bit] Compiling...
cl /LD /W4 /nologo /utf-8 "%SRC_DIR%TextTools.c" /link /DEF:"%SRC_DIR%TextTools.def" /OUT:"%SRC_DIR%TextTools64.dll" /MACHINE:X64
if %ERRORLEVEL% NEQ 0 (
    echo [64bit] FAILED with error %ERRORLEVEL%
    goto :done
)
echo [64bit] OK

REM ---- 32bit ----
echo [32bit] Setting up environment...
call %VCVARS% x86 > nul 2>&1
echo [32bit] Compiling...
cl /LD /W4 /nologo /utf-8 "%SRC_DIR%TextTools.c" /link /DEF:"%SRC_DIR%TextTools.def" /OUT:"%SRC_DIR%TextTools32.dll" /MACHINE:X86
if %ERRORLEVEL% NEQ 0 (
    echo [32bit] FAILED with error %ERRORLEVEL%
    goto :done
)
echo [32bit] OK

:done
echo === Build finished ===
endlocal
