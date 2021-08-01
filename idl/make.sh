#!/bin/bash
set echo off
# Linux make-file for Interpolation LOo Calc Add-in.
# Created by jan@biochemfusion.com, April 2009 for Doobidoo Add In
# Adapted by manuel_astros@outlook.com, July 2018

#SET LOO_HOME=C:\Program Files\LibreOffice
LOO_HOME=/opt/libreoffice6.4

LOO_BIN_DIR=$LOO_HOME/program
TOOLS_BIN_DIR=$LOO_HOME/sdk/bin

PACKAGE_NAME=Interpolation

# The IDL tools rely on supporting files in the main OOo installation.
PATH=$PATH:$LOO_HOME/program

#
# Compile IDL file.
#
#cp /windows/INTERPOLACION-APP/'Interpolation - Py'/idl/XInterpolation.idl %LOO_HOME%/sdk/idl
#SET IDL_INCLUDE_DIR=%LOO_HOME%/sdk/bin/idl
IDL_INCLUDE_DIR=/home/manuel/Documents/LOO-Develop
IDL_FILE=X$PACKAGE_NAME
PATH=$PATH:/opt/libreoffice6.4/sdk/bin

cd $TOOLS_BIN_DIR
idlc -w -C -I "../idl" "$IDL_INCLUDE_DIR/$IDL_FILE.idl"

# Convert compiled IDL to loadable type library file.
# First remove existing .rdb file, otherwise regmerge will just
# append the compiled IDL to the resulting .rdb file. The joy of
# having an .rdb file with several conflicting versions of compiled
# IDL is very very limited - don't go there.

PATH=$PATH:$TOOLS_BIN_DIR
if [ -f $IDL_FILE.rdb ]
then
	rm $IDL_FILE.rdb
else
	cd $IDL_INCLUDE_DIR
	"$LOO_BIN_DIR/regmerge" $IDL_FILE.rdb /UCR $IDL_FILE.urd
fi
rm $IDL_FILE.urd
