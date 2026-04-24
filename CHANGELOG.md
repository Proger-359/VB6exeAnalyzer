# Changelog

## 

## 2026-04-24

All found backups and distribution package have been published for archival and educational use.

#### Changed :

- All source files have been moved into proper subfolders : classe, modules, forms.
- Main VBP file has been updated accordingly.
- Tools sub-project files have been moved to their own subfolders. 
- Compilated and executables files have been removed.

---

## 2011-11-17 09:29pm

Last known update on this project

---

###### Following changelog is adapted from the original historical update log (available in french at `docs/`).

## 2008-11-17

### Fixed

- Additional crash fixes for large/complex executables.

- Minor improvements.

### Added

- Additional missing tools (e.g., StrRef right-click menu).

- More complete string and object reference search.

---

## 2008-11-17 11:23am

#### Fixed

- Various crash fixes when disassembling complex executables.

#### Improved

- More complete search for string references and object references.

#### Added

- Some more tools:
  
  - Using Right-click in StrRef list now pops a menu

---

## 2006-06-14 00:52am

#### Added

- New tool: disassembly using **PERDR**
  
  - Much more precise but slower than the built-in routine.
  
  - Note :
    
    - Download PERDR from:  
      http://sourceforge.net/projects/perdr
    
    - Place `perdr.exe` in the project directory.
    
    - Launch the project, it will automatically use it.

#### Improved

- Interpreter:
  
  - Much more accurate detection of control calls and properties.
  
  - Foundation for reconstructing code like:
    
    ```
    form1.label5.caption = "toto"
    form2.command2.enabled = True
    ```

### Changed

- Some code refactor and bug fixes.

---

## 2006-01-22 01:54am

#### Added

- Analyzer now pre-scans compiled code :
  
  - Finds additional Subs (including `.bas` module Subs).

- Interpreter:
  
  - Detects Sub calls (`Call Sub x`).
  
  - Finds object references (not yet identified precisely).

- Interface:
  
  - Progress indicators during disassembly.
  
  - Two additional internal tools.
  
  - Two root RVA values added to hierarchical report.

#### Fixed

- No longer crashes when a `Sub Main` exists.

---

## 2006-01-19 11:07am

#### Added

- Interpreter now integrates **EP (Entry Points)** : useful for understanding object calls.

#### Fixed

- 3 disassembler instruction fixes (including negative `dword ptr`).

- Analyzer: object Sub recovery fixes.

- Hierarchical report improvements.

- Several crash fixes.

---

## 2005-12-27 00:14am

#### Fixed

- Support for bigger, complex VB6 executables.

#### Changed

- Main form redesigned: all features accessible via menus.

#### Added

- Export analysis results to text files.

- Useful for continuing the project or building new interfaces.

#### Known issues

- Conditional jumps are often incorrect.

- Disassembler becomes difficult to debug; may need full rewrite.

#### Improved

- Sub display in interpreter and hierarchical report.

- Various small tweaks.

---

## 2005-12-24 06:31pm

#### Fixed

- Correct handling of 16-bit prefix used by VB with Integers.

#### Improved

- Interpreter algorithm optimized (**~50% faster**).

#### Added

- Detection of declared API calls (`Declare Sub`) in disassembled code.

- Ability to list found strings in a separate list box.

---

## 2005-12-21 05:39pm

#### Changed

- New Sub analysis:
  
  - Much more accurate.
  
  - Associates Subs with forms and objects.
  
  - Object hex value → name mapping still missing.

#### Fixed

- Major module analysis cleanup after re-reading the code.

- Removed large amount of debug code.

---

## 2005-12-11 11:39am

#### Added

- 4 new correctly disassembled instructions.

#### Improved

- Faster disassembly.

- Better string recovery.

- Listing search capability.

- Jump/call tracking.

- Sub tracking.

---

## 2005-12-04 04:41pm

#### Fixed

- Better handling of operand-size prefix and common instruction parsing.

#### Improved

- Interpreter now detects API calls and strings more exhaustively.

#### Added

- Disassembler feature: **Follow RVA** to read instruction targets.

---

## 2005-12-02 07:20pm

#### Fixed

- Disassembler fixes for some incorrectly interpreted standard instructions.

#### Improved

- Better string recovery (strings as in-code vars, not properties).

- Improved hex preview tool.

---

## 2005-12-02 00:28am

#### Fixed

- Disassembler fixes:
  
  - Some FPU instructions.
  
  - Display order issues.

#### Added

- Disassembler now retrieves **string calls**.

---

## 2005-12-01 05:23pm

#### Changed

- Main interface reduced to essential features.

#### Improved

- Major improvement and debugging of the disassembler.
  
  - About **85% of the code is now correctly disassembled** (vs 1% previously).

- Improved presentation of disassembled code.

- New interface for the disassembler/interpreter.

- Interpreter now detects API calls from imported PE DLLs (mainly `msvbm60.dll`).

- Reworked Sub table analysis.
  
  - All calls are now detected, but not yet linked to forms/objects/modules or categorized as Sub/Function/Event.

#### Removed

- Recreation of empty sources files with known objects properties

---

## 2005-01-31 01:26pm

#### Added

- Integrated `memwork.cls` and the original tool that uses it.

---

###### Following changelog is adapted from misc comments found either in code or from archive.org [backup of comments](https://web.archive.org/web/20070213010640/http://www.vbfrance.com/codes/ANALYSEUR-EXECUTABLES-VB6-BETA-RELEASE_8258.aspx)

## 2004-02-29 11:00pm

#### Fixed

- Support for invisible item menu string

---

## 2003-09-23 08:00pm

#### Added

- Now parse Subs

---

## 2003-08-17 12:00pm

Initial public release
