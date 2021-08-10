#!/bin/bash

cd ..

python3 -m venv env

source env/bin/activate

mkdir 'excel files'

pip3 install openpyxl
pip3 install xls2xlsx

$SHELL
