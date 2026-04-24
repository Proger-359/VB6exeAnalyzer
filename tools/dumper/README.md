# Dumper

This is a historical reverse engineering project from the VB6 era, released for educational and archival purposes.

"Dumper" is a side tool for the VB6 EXE Analyzer project.



## Overview

Originally, a tool to imitate Cheat-o-Matic : a program that hacked into memory space of running videogame in order to change realtime in-game values (like health point, money, or experience points)

Later, a way to dump the content of any program loaded in memory into a cold EXE file. That ability allowed to get unpacked content of executable files.



## Usage

You will need administrator privilege in order to access other running processes outside dumper.exe itself

1. Launch the program (running from VB6 IDE works)

2. Select a running process from the combo box

3. Click on the wanted module (loaded binaries in memory, usually the main EXE and many DLLs)

4. Click on "Dumper" button from within the "Info module" tab to get an hex view window with the raw memory content.

The program could fail due to antivirus software or limited permissions.



## License

This project is released under the MIT License.


