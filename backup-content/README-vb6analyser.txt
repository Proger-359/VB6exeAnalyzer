==============
 VB6 ANALYZER
==============

VB6's exe files parsing, disassembling and half decompilation

by Proger, 2003-2006
last update nov.2008

mailto:proger@cbsky.net
(this mailbox is spammed, so use relevant subject if you want me to find your mail)


PROJECT SUMMARY

first goal is to check how the VB compilator translate the high-level code to machine code,
 in order to make very efficient subs and algorithms.
indirects goals was set, since i needed to recover many datas in order to find accurately where
 is my algorithm inside that 'mess'. Moreover, the more automatic description i can add near
 the disassembled code list, the easier you and I will understand what happens in vb6 mind :)


PURPOSE OF THIS FILE

For you to understand inside code of the project.
All the comments in the project are in french, thus you'll need this file.


OVERVIEW OF PROJECT FILES

PEexe.bas : main module of PE exe reading, VB6 exe data parsing and reporting
unASM.bas : module of disassembly code, and VB6 native code disassembling
Analyse.bas : module trying to decompile the disassembled vb listing (incomplete but enough for me)
VbObjAnalyse.bas : module to analyse properties/call for each VB objects (unfinished, i'm lazy)
PeReaderExt.bas : module that use perdr.exe to do the diassembly job (because my disassembling algo sux at FPU instructions)

frmPeExe.frm : main form, code for file output of analysis in menu
frmList.frm : pure listbox in a form, multiple purpose
frmResultat.frm : pure textbox in a form, multiple purpose
frmHexa.frm : hex view form, with two tools embedded
frmJolieInterface.frm : a treeview with some behavior on node-click. Used for pretty reporting
frmUnAsm.frm : a listview, used for unassembly result print, and little tools embedded in

VB60_APIDEF.txt : mandatory file with all the msvbmv60.dll export API table.


OVERVIEW OF MAIN SUBS

PEexe.OpenEXE() : Starting sub for analysis. It parses PE data then chain with VB6 parsing functions
> PEexe.FindControl() : Search and parse the forms include in the VB6 exe
-> PEexe.ParseControl() : Search the objects in any forms found
-> PEexe.ParseControlParams() : Recover some objetcs attributes (buggy and not complete since others project make it well better)
> PEexe.FindModules() : Search all the modules (forms,modules,classes,userctls) and find names, subs and user API declare. Also find native code start offset.
-> PEexe.ParseVBSubs() : pinpoint subs entry point in native code
-> PEexe.ParseCode() : first scan of the native code to find 'not listed' subs (e.g subs from .bas modules)
> PEexe.ParseVBEP() : parse the structure table who embed various pointers inside subs in native code. Usefulness not yet indentified.
> PEexe.AssocieSubObj() : try to find the exact sub origin (objects, event names...)


unASM.VBCODE_DeAsm() : Starting sub of disassembly and VB6 pre-decompilation engine
-- first part is building a list with all the known relevant offsets to api/ep/subs and descriptions
-- second part is byte-by-byte disassembly
TblPtrAsm() is an index table to quickly find the closest instruction matching the byte read, then it chains to...
> unASM.GetVASM() : search the correct instruction matching the word (opcode) read ; a simple seek engine
> unASM.CodeToStr() : the disassembler. Many tuning with 'wobbly' bugfixes forbid me to explain how it works in one sentence
-- in the disassembly loop, there is various scanning, attempting to recover strings references, objects references, and matching with relevant offsets.
-- THIS LAST SUB NEED TO BE IMPROVED FOR FPU opcode.

