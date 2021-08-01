@echo off
:: Windows make-file for Interpolation LOo Calc Add-in.
:: Created by jan@biochemfusion.com, April 2009 for Doobidoo Add In
:: Adapted by manuel_astros@outlook.com, July 2018

SET LOO_HOME=C:\Program Files\LibreOffice

SET LOO_BIN_DIR=%LOO_HOME%\program
SET TOOLS_BIN_DIR=%LOO_HOME%\sdk\bin

SET PACKAGE_NAME=Interpolation

:: The IDL tools rely on supporting files in the main OOo installation.
PATH=%PATH%;%LOO_HOME%\program

::
:: Compile IDL file.
::

SET IDL_INCLUDE_DIR=%LOO_HOME%\sdk\bin\idl
SET IDL_FILE=X%PACKAGE_NAME%

cd %TOOLS_BIN_DIR%
idlc.exe -w -C -I "..\idl" "%IDL_INCLUDE_DIR%\%IDL_FILE%.idl"

:: Convert compiled IDL to loadable type library file.
:: First remove existing .rdb file, otherwise regmerge will just
:: append the compiled IDL to the resulting .rdb file. The joy of
:: having an .rdb file with several conflicting versions of compiled
:: IDL is very very limited - don't go there.
if exist %IDL_FILE%.rdb. (
del %IDL_FILE%.rdb
)
cd %IDL_INCLUDE_DIR%
"%LOO_BIN_DIR%\regmerge.exe" %IDL_FILE%.rdb /UCR %IDL_FILE%.urd

del %IDL_FILE%.urd


