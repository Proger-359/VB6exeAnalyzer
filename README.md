# VB6 EXE Analyzer

This is a historical reverse engineering project from the VB6 era, released for educational and archival purposes.

## Overview

A Visual Basic 6 native code disassembler written itself in VB6, originally developed around 2003-2008 and released with full source code.

## Features

- Disassembling and partial interpretation of VB6 executable files (native code).
- Can unpack PE packed files with UPX and alike before parsing.
- View screen with following details:
  - Source files : modules, forms, classes, userctls, with their names.
  - Function in each file (with code entry point offset).
  - Objects in UI (frm), events function (code entry point offset) and properties.
  - External api calls (p/invoke).
- View screen with disassembled code:
  - Ordered by functions offsets.
  - Partial interpretation.
  - Follow rva (relative virtual address jump), string search.
- Export analysis result to various textfiles.
- Internal disassembler, pure VB6 code.
- optionaly use an external disassembler (PEReaDeR) for better FPU opcode disassembly.
- Include exported API of msvbm60.dll with partial interpretation.

## Screenshots

Original screenshot (`/screenshots/demo.jpg`)

![Main UI](D:\FTP\TI\vb6-legacy\8258\repo\screenshots\demo.jpg)

## Getting started

### Requirement

- Visual Basic 6 IDE (for source compilation or to run in the IDE).

### Usage

1. Starts with menu "Main" > "Ouvrir un executable" , then select an executable file.

2. Menu "Main" > "Analyser maintenant!".

3. Now all other menu items are enabled.

4. For the disassembler view : Menu "Analyse" > "Listing désassemblé" > answer "Yes" (or "Oui") to launch the disassembly that take a few seconds.

5. For the treeview with all elements found : "Analyse" > "Rapport hiérarchisé".

## Limitations

- Internal disassembler doesn't support well the FPU Opcode.

- Partial interpretation using msvbmv60.dll API calls translation.

- Only a few objects properties are retrieved.

- Does not match UI objects function with their events.

- Could require administrator privilege to unpack Exe-packed files.

## Project structure

- `src/` – VB6 source code.
- `docs/` – technical documentation  and historical notes.
- `tools/`- misc VB6 programs used to create or test part of the main project.

## Historical context

This project was originally released on www.vbfrance.com (website before codes-sources.com) around 2003 (link to archive.org [here]([ANALYSEUR D'EXECUTABLES VB6 (BETA RELEASE) désassembleur, analyse, executable, décompilateur, analyseur ☼ Code source N°8258 ☼ Visual Basic, VB6, VB.NET, VB 2005](https://web.archive.org/web/20070213010640/http://www.vbfrance.com/codes/ANALYSEUR-EXECUTABLES-VB6-BETA-RELEASE_8258.aspx)))

Only VBDE and patched Win32DAsm were able to correctly parse VB6 compilated binaries in native mode when this project started.

## Roadmap

- Finding UI objects event functions addresses.

- Better interpretation and FPU Opcode support.

- Finding and printing full dependencies (external OCX and DLL).

- Restore native-exe to empty source files rebuilder (removed in 2005).

## License

This project is released under the MIT License.
