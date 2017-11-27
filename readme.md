# Windows - Script utilities for Windows & MS Office

> See also my [MS Access repository](https://github.com/cavo789/ms_access) which contains vba code for MS Access (to place in a module)

This repository contains VBS classes that will help Windows users to collect a list of files based on extensions (see `classes/ScanFolder.vbs`) and utilities to work with MS Access databases (see `classes/MSAccess.vbs`) or MS Excel workbooks (see `classes\MSExcel.vbs`).

## classes/Files.vbs

Provide functionnalities for working with files under Windows.

See [documentation](src/classes/files.md)

## classes/MSAccess.vbs

Provide functionnalities for working with MS Access databases.

See [documentation](src/classes/msaccess.md)

## classes/MSExcel.vbs

Provide functionnalities for working with MS Excel workbooks.

See [documentation](src/classes/msexcel.md)

## classes/ScanFolder.vbs

ScanFolder is aiming to scan recursively a folder and search for f.i. text files (`.txt`, `.csv`, `.md`, ...), Excel files (`.xlsx`, `.xlam`, ...), collect the list of filenames and output them in a text file so that file can be used by an another business logic.

See [documentation](src/classes/scanfolder.md)
