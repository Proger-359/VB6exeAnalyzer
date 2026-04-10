Work in progress ...

PRE-DECOMPILING :
identifying native code compiled EXE using msvbm60.dll
retrieve program entry point
retrieve texts entry point
retrieve controls, forms, ..., entry point

STEP 1 :
identifying property of each common control of VB6

POST STEP 1 :
creating frm file with each control and control property of one VB PE EXE

STEP 2 :
identifying hex code of common vb's function
If Then Else Goto End For Next Exit Do Loop
identifying parameters

identifying hex code of extended  vb's function
Asc Chr Len Mid Trim, operators, &, ...
these functions calls API beginning with __vba

POST STEP 2 :
rebuild partial code in each control used in a form

STEP 3 :
identifying use of module or class an rebuilding code, subs and calls

