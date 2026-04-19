Attribute VB_Name = "UnASMv2"
'=========================
' UnASM v2
'**********
'
'par Proger
'création octobre 2005
'
'
'désassembleur de code machine vers listing assembleur
'version 2 : tri par arbre

Option Explicit
DefLng A-Z

Private Type ASM_INSTRUCT
    aStart As Byte        'octet d'entête signalant l'instruction
    bPrefix As Boolean    'indique qu'il s'agit d'un prefix
    cUseInt As Boolean    'signale si l'entête à deux octet
    dStart2 As Byte       'deuxième octet signalant l'instruction
    eDirect As Boolean    'indique que l'instruction n'as aucuns paramètres
    
    fUseModRM As Boolean  'indique que l'octet suivant est ModR/M
    gRMtype As Byte       'mode de R/M : 8-bit, 16-bit, 32-bit
    hUseRegEx As Boolean  'indique que la partie Reg1 de ModR/M permet de distinguer l'instruction
    iRegExVal As Byte     'valeur de la partie Reg1 (en cas de UseRegEx = True) de cette instruction (0-7)
    
    kSpecial As Boolean   'indique que l'instruction nécessite un traitement particulier
    
    lUseImm As Boolean    'indique qu'il y a utilisation d'une valeure immédiate
    mImmtype As Byte      'indique le mode de l'immédiat : 8,16,32,64-bits...
    
    nUseRel As Boolean    'indique qu'il y a utilisation d'un saut relatif
    oReltype As Byte      'indique le mode de saut : 8,16,32-bits
    
    pUseOffset As Boolean
    
    xAdv As Boolean       'indique s'il s'agit d'une instruction spéciale (mmx/3dnow/sse/...)
    
    zInstruct As String   'texte de début de l'instruction
    'format d'affichage :
    '%i = immediat
    '%r = registre
    '%rm = registre ou valeur (modr/m)
    '%j = addresse de saut, a partir d'une adresse relative
    '%o = offset (?)
    
    
End Type
    Private tbl_ASM() As ASM_INSTRUCT
    Private tbl_ASMlen As Long


Private Type ASM_REGISTER
    r8 As String * 2
    r16 As String * 2
    r32 As String * 3
    r64 As String * 3
End Type
    Private TblASM_REG(0 To 7) As ASM_REGISTER

Sub UnAsm_Init(Optional IA32 As Boolean = True, _
               Optional x87fpu As Boolean = True, _
               Optional x87fpu_simd As Boolean = False, _
               Optional MMX As Boolean = False, _
               Optional SSE As Boolean = False, _
               Optional SSE2 As Boolean = False, _
               Optional SSE3 As Boolean = False, _
               Optional IA32e As Boolean = False)
               
    'registres
    TblASM_REG(0).r8 = "al": TblASM_REG(0).r16 = "ax": TblASM_REG(0).r32 = "eax": TblASM_REG(0).r64 = "rax"
    TblASM_REG(1).r8 = "cl": TblASM_REG(1).r16 = "cx": TblASM_REG(1).r32 = "ecx": TblASM_REG(1).r64 = "rcx"
    TblASM_REG(2).r8 = "dl": TblASM_REG(2).r16 = "dx": TblASM_REG(2).r32 = "edx": TblASM_REG(2).r64 = "rdx"
    TblASM_REG(3).r8 = "bl": TblASM_REG(3).r16 = "bx": TblASM_REG(3).r32 = "ebx": TblASM_REG(3).r64 = "rbx"
    TblASM_REG(4).r8 = "ah": TblASM_REG(4).r16 = "sp": TblASM_REG(4).r32 = "esp": TblASM_REG(4).r64 = "rsp"
    TblASM_REG(5).r8 = "ch": TblASM_REG(5).r16 = "bp": TblASM_REG(5).r32 = "ebp": TblASM_REG(5).r64 = "rbp"
    TblASM_REG(6).r8 = "dh": TblASM_REG(6).r16 = "si": TblASM_REG(6).r32 = "esi": TblASM_REG(6).r64 = "rsi"
    TblASM_REG(7).r8 = "bh": TblASM_REG(7).r16 = "di": TblASM_REG(7).r32 = "edi": TblASM_REG(7).r64 = "rdi"
    
    'instructions
    ReDim tbl_ASM(1 To 1500) As ASM_INSTRUCT
    tbl_ASMlen = 0
    
    If IA32 Then
        Call Instr_IA32
    End If
    
    If x87fpu Then
        Call Instr_x87fpu
    End If
    
    If x87fpu_simd Then
        Call Instr_x87fpu_simd
    End If
    
    If MMX Then
        Call Instr_MMX
    End If
    
    If SSE Then
        Call Instr_SSE
    End If
    If SSE2 Then
        Call Instr_SSE2
    End If
    If SSE3 Then
        Call Instr_SSE3
    End If
    If IA32e Then
        Call Instr_IA32e
    End If

End Sub

Private Sub Instr_IA32()
'instructions standard intel pour 32-bit "x386" : General Instructions

'part 1 : Dara Transfer Instructions
'
'MOV Move data between general-purpose registers; move data between memory and general-purpose or segment registers; move immediates to general-purpose registers
'CMOVE/CMOVZ Conditional move if equal/Conditional move if zero
'CMOVNE/CMOVNZ Conditional move if not equal/Conditional move if not zero
'CMOVA/CMOVNBE Conditional move if above/Conditional move if not below or equal
'CMOVAE/CMOVNB Conditional move if above or equal/Conditional move if not below
'CMOVB/CMOVNAE Conditional move if below/Conditional move if not above or equal
'CMOVBE/CMOVNA Conditional move if below or equal/Conditional move if not above
'CMOVG/CMOVNLE Conditional move if greater/Conditional move if not less or equal
'CMOVGE/CMOVNL Conditional move if greater or equal/Conditional move if not less
'CMOVL/CMOVNGE Conditional move if less/Conditional move if not greater or equal
'CMOVLE/CMOVNG Conditional move if less or equal/Conditional move if not greater
'CMOVC Conditional move if carry
'CMOVNC Conditional move if not carry
'CMOVO Conditional move if overflow
'CMOVNO Conditional move if not overflow
'CMOVS Conditional move if sign (negative)
'CMOVNS Conditional move if not sign (non-negative)
'CMOVP/CMOVPE Conditional move if parity/Conditional move if parity even
'CMOVNP/CMOVPO Conditional move if not parity/Conditional move if parity odd
'XCHG exchange
'BSWAP Byte swap
'XADD exchange And Add
'CMPXCHG Compare And exchange
'CMPXCHG8B Compare and exchange 8 bytes
'PUSH Push onto stack
'POP Pop off of stack
'PUSHA/PUSHAD Push general-purpose registers onto stack
'POPA/POPAD Pop general-purpose registers from stack
'CWD/CDQ Convert word to doubleword/Convert doubleword to quadword
'CBW/CWDE Convert byte to word/Convert word to doubleword in EAX register
'MOVSX Move and sign extend
'MOVZX Move and zero extend

'part 2 : Binary Arithmetic Instructions
'ADD Integer add
'ADC Add with carry
'SUB Subtract
'SBB Subtract with borrow
'IMUL Signed multiply
'MUL Unsigned multiply
'IDIV Signed divide
'DIV Unsigned divide
'INC Increment
'DEC Decrement
'NEG Negate
'CMP Compare

'part 3 : Decimal Arithmetic Instructions (sur données BCD)
'DAA Decimal adjust after addition
'DAS Decimal adjust after subtraction
'AAA ASCII adjust after addition
'AAS ASCII adjust after subtraction
'AAM ASCII adjust after multiplication
'AAD ASCII adjust before division

'part 4 : Logical Instructions
'AND Perform bitwise logical AND
'OR Perform bitwise logical OR
'XOR Perform bitwise logical exclusive OR
'NOT Perform bitwise logical NOT

'part 5 : Shift and Rotate Instructions
'SAR Shift arithmetic right
'SHR Shift logical right
'SAL/SHL Shift arithmetic left/Shift logical left
'SHRD Shift right double
'SHLD Shift left double
'ROR Rotate right
'ROL Rotate left
'RCR Rotate through carry right
'RCL Rotate through carry left

'part 6 : Bit and Byte Instructions - utilise EFLAGS
'BT Bit test
'BTS Bit test and set
'BTR Bit test and reset
'BTC Bit test and complement
'BSF Bit scan forward
'BSR Bit scan reverse
'SETE/SETZ Set byte if equal/Set byte if zero
'SETNE/SETNZ Set byte if not equal/Set byte if not zero
'SETA/SETNBE Set byte if above/Set byte if not below or equal
'SETAE/SETNB/SETNC Set byte if above or equal/Set byte if not below/Set byte if not carry
'SETB/SETNAE/SETC Set byte if below/Set byte if not above or equal/Set byte if carry
'SETBE/SETNA Set byte if below or equal/Set byte if not above
'SETG/SETNLE Set byte if greater/Set byte if not less or equal
'SETGE/SETNL Set byte if greater or equal/Set byte if not less
'SETL/SETNGE Set byte if less/Set byte if not greater or equal
'SETLE/SETNG Set byte if less or equal/Set byte if not greater
'SETS Set byte if sign (negative)
'SETNS Set byte if not sign (non-negative)
'SETO Set byte if overflow
'SETNO Set byte if not overflow
'SETPE/SETP Set byte if parity even/Set byte if parity
'SETPO/SETNP Set byte if parity odd/Set byte if not parity
'TEST Logical compare

'part 7 : Control Transfer Instructions
'JMP Jump
'JE/JZ Jump if equal/Jump if zero
'JNE/JNZ Jump if not equal/Jump if not zero
'JA/JNBE Jump if above/Jump if not below or equal
'JAE/JNB Jump if above or equal/Jump if not below
'JB/JNAE Jump if below/Jump if not above or equal
'JBE/JNA Jump if below or equal/Jump if not above
'JG/JNLE Jump if greater/Jump if not less or equal
'JGE/JNL Jump if greater or equal/Jump if not less
'JL/JNGE Jump if less/Jump if not greater or equal
'JLE/JNG Jump if less or equal/Jump if not greater
'JC Jump if carry
'JNC Jump if not carry
'JO Jump if overflow
'JNO Jump if not overflow
'JS Jump if sign (negative)
'JNS Jump if not sign (non-negative)
'JPO/JNP Jump if parity odd/Jump if not parity
'JPE/JP Jump if parity even/Jump if parity
'JCXZ/JECXZ Jump register CX zero/Jump register ECX zero
'LOOP Loop with ECX counter
'LOOPZ/LOOPE Loop with ECX and zero/Loop with ECX and equal
'LOOPNZ/LOOPNE Loop with ECX and not zero/Loop with ECX and not equal
'CALL Call procedure
'RET Return
'IRET Return from interrupt
'INT Software interrupt
'INTO Interrupt on overflow
'BOUND Detect value out of range
'ENTER High-level procedure entry
'LEAVE High-level procedure exit

'part 8 : String Instructions
'MOVS/MOVSB Move string/Move byte string
'MOVS/MOVSW Move string/Move word string
'MOVS/MOVSD Move string/Move doubleword string
'CMPS/CMPSB Compare string/Compare byte string
'CMPS/CMPSW Compare string/Compare word string
'CMPS/CMPSD Compare string/Compare doubleword string
'SCAS/SCASB Scan string/Scan byte string
'SCAS/SCASW Scan string/Scan word string
'SCAS/SCASD Scan string/Scan doubleword string
'LODS/LODSB Load string/Load byte string
'LODS/LODSW Load string/Load word string
'LODS/LODSD Load string/Load doubleword string
'STOS/STOSB Store string/Store byte string
'STOS/STOSW Store string/Store word string
'STOS/STOSD Store string/Store doubleword string
'REP Repeat while ECX not zero
'REPE/REPZ Repeat while equal/Repeat while zero
'REPNE/REPNZ Repeat while not equal/Repeat while not zero

'part 9 : I/O Instructions
'IN Read from a port
'OUT Write to a port
'INS/INSB Input string from port/Input byte string from port
'INS/INSW Input string from port/Input word string from port
'INS/INSD Input string from port/Input doubleword string from port
'OUTS/OUTSB Output string to port/Output byte string to port
'OUTS/OUTSW Output string to port/Output word string to port
'OUTS/OUTSD Output string to port/Output doubleword string to port

'part 10 : c'est ENTER et LEAVE

'part 11 : Flag Control (EFLAG) Instructions
'STC Set carry flag
'CLC Clear the carry flag
'CMC Complement the carry flag
'CLD Clear the direction flag
'STD Set direction flag
'LAHF Load flags into AH register
'SAHF Store AH register into flags
'PUSHF/PUSHFD Push EFLAGS onto stack
'POPF/POPFD Pop EFLAGS from stack
'STI Set interrupt flag
'CLI Clear the interrupt flag

'part 12 : Segment Register Instructions
'LDS Load far pointer using DS
'LES Load far pointer using ES
'LFS Load far pointer using FS
'LGS Load far pointer using GS
'LSS Load far pointer using SS

'part 13 : Miscellaneous Instructions
'LEA Load effective address
'NOP No operation
'UD2 Undefined instruction <== mort de rire :)
'XLAT/XLATB Table lookup translation
'CPUID Processor Identification

'extra : SYSTEM INSTRUCTIONS
'instructions spéciale pour OS... en gros c'est pour les interruptions et le multitâche.
'LGDT Load global descriptor table (GDT) register
'SGDT Store global descriptor table (GDT) register
'LLDT Load local descriptor table (LDT) register
'SLDT Store local descriptor table (LDT) register
'LTR Load task register
'STR Store task register
'LIDT Load interrupt descriptor table (IDT) register
'SIDT Store interrupt descriptor table (IDT) register
'MOV Load and store control registers
'LMSW Load machine status word
'SMSW Store machine status word
'CLTS Clear the task-switched flag
'ARPL Adjust requested privilege level
'LAR Load access rights
'LSL Load segment limit
'VERR Verify segment for reading
'VERW Verify segment for writing
'MOV Load and store debug registers
'INVD Invalidate cache, no writeback
'WBINVD Invalidate cache, with writeback
'INVLPG Invalidate TLB Entry
'LOCK (prefix) Lock Bus
'HLT Halt processor
'RSM Return from system management mode (SMM)
'RDMSR Read model-specific register
'WRMSR Write model-specific register
'RDPMC Read performance monitoring counters
'RDTSC Read time stamp counter
'SYSENTER Fast System Call, transfers to a flat protected mode kernel at CPL = 0
'SYSEXIT Fast System Call, transfers to a flat protected mode kernel at CPL = 3


End Sub
Private Sub Instr_x87fpu()
'instructions utilisé depuis le 80486 : opérations sur floating point et binary coded decimal (BCD)

'part 1 : x87 FPU Data Transfer Instructions
'FLD Load floating-point value
'FST Store floating-point value
'FSTP Store floating-point value and pop
'FILD Load integer
'FIST Store integer
'FISTP Store integer and pop (n'existe que en SSE3)
'FBLD Load BCD
'FBSTP Store BCD and pop
'FXCH Exchange registers
'FCMOVE Floating-point conditional move if equal
'FCMOVNE Floating-point conditional move if not equal
'FCMOVB Floating-point conditional move if below
'FCMOVBE Floating-point conditional move if below or equal
'FCMOVNB Floating-point conditional move if not below
'FCMOVNBE Floating-point conditional move if not below or equal
'FCMOVU Floating-point conditional move if unordered
'FCMOVNU Floating-point conditional move if not unordered

'part 2 : x87 FPU Basic Arithmetic Instructions
'FADD Add floating-point
'FADDP Add floating-point and pop
'FIADD Add integer
'FSUB Subtract floating-point
'FSUBP Subtract floating-point and pop
'FISUB Subtract integer
'FSUBR Subtract floating-point reverse
'FSUBRP Subtract floating-point reverse and pop
'FISUBR Subtract integer reverse
'FMUL Multiply floating-point
'FMULP Multiply floating-point and pop
'FIMUL Multiply integer
'FDIV Divide floating-point
'FDIVP Divide floating-point and pop
'FIDIV Divide integer
'FDIVR Divide floating-point reverse
'FDIVRP Divide floating-point reverse and pop
'FIDIVR Divide integer reverse
'FPREM Partial remainder
'FPREM1 IEEE Partial remainder
'FABS Absolute value
'FCHS Change sign
'FRNDINT Round to integer
'FSCALE Scale by power of two
'FSQRT Square root
'FXTRACT Extract exponent and significand

'part 3 : x87 FPU Comparison Instructions
'FCOM Compare floating-point
'FCOMP Compare floating-point and pop
'FCOMPP Compare floating-point and pop twice
'FUCOM Unordered compare floating-point
'FUCOMP Unordered compare floating-point and pop
'FUCOMPP Unordered compare floating-point and pop twice
'FICOM Compare integer
'FICOMP Compare integer and pop
'FCOMI Compare floating-point and set EFLAGS
'FUCOMI Unordered compare floating-point and set EFLAGS
'FCOMIP Compare floating-point, set EFLAGS, and pop
'FUCOMIP Unordered compare floating-point, set EFLAGS, and pop
'FTST Test floating-point (compare with 0.0)
'FXAM Examine floating-point

'part 4 : x87 FPU Transcendental Instructions (trigo basique et loga)
'FSIN Sine
'FCOS cosine
'FSINCOS Sine and cosine
'FPTAN Partial tangent
'FPATAN Partial arctangent
'F2XM1 2^x - 1
'FYL2X y * log2(x)
'FYL2XP1 y * log2(x + 1)

'part 5 : x87 FPU Load Constants Instructions
'FLD1 Load + 1.0
'FLDZ Load + 0.0
'FLDPI Load Pi
'FLDL2E Load log2(e)
'FLDLN2 Load loge(2)
'FLDL2T Load log2(10)
'FLDLG2 Load log10(2)

'part 6 : x87 FPU Control Instructions
'FINCSTP Increment FPU register stack pointer
'FDECSTP Decrement FPU register stack pointer
'FFREE Free floating-point register
'FINIT Initialize FPU after checking error conditions
'FNINIT Initialize FPU without checking error conditions
'FCLEX Clear floating-point exception flags after checking for error conditions
'FNCLEX Clear floating-point exception flags without checking for error conditions
'FSTCW Store FPU control word after checking error conditions
'FNSTCW Store FPU control word without checking error conditions
'FLDCW Load FPU control word
'FSTENV Store FPU environment after checking error conditions
'FNSTENV Store FPU environment without checking error conditions
'FLDENV Load FPU environment
'FSAVE Save FPU state after checking error conditions
'FNSAVE Save FPU state without checking error conditions
'FRSTOR Restore FPU state
'FSTSW Store FPU status word after checking error conditions
'FNSTSW Store FPU status word without checking error conditions
'WAIT/FWAIT Wait for FPU
'FNOP FPU no operation


End Sub
Private Sub Instr_x87fpu_simd()
'instructions utilisé à partir du Pentium 2, SIMD (single-instruction multiple-data) state management

'FXSAVE Save x87 FPU and SIMD state
'FXRSTOR Restore x87 FPU and SIMD state


End Sub
Private Sub Instr_MMX()
'instructions spéciale Pentium MMX : gestion des entiers par paquets en SIMD
'pourquoi MMX ? parce que le registre 64-bits s'appelle MMX, lol

'part 1 : MMX Data Transfer Instructions
'MOVD Move doubleword
'MOVQ Move quadword

'part 2 : MMX Conversion Instructions
'PACKSSWB Pack words into bytes with signed saturation
'PACKSSDW Pack doublewords into words with signed saturation
'PACKUSWB Pack words into bytes with unsigned saturation.
'PUNPCKHBW Unpack high-order bytes
'PUNPCKHWD Unpack high-order words
'PUNPCKHDQ Unpack high-order doublewords
'PUNPCKLBW Unpack low-order bytes
'PUNPCKLWD Unpack low-order words
'PUNPCKLDQ Unpack low-order doublewords

'part 3 : MMX Packed Arithmetic Instructions
'PADDB Add packed byte integers
'PADDW Add packed word integers
'PADDD Add packed doubleword integers
'PADDSB Add packed signed byte integers with signed saturation
'PADDSW Add packed signed word integers with signed saturation
'PADDUSB Add packed unsigned byte integers with unsigned saturation
'PADDUSW Add packed unsigned word integers with unsigned saturation
'PSUBB Subtract packed byte integers
'PSUBW Subtract packed word integers
'PSUBD Subtract packed doubleword integers
'PSUBSB Subtract packed signed byte integers with signed saturation
'PSUBSW Subtract packed signed word integers with signed saturation
'PSUBUSB Subtract packed unsigned byte integers with unsigned saturation
'PSUBUSW Subtract packed unsigned word integers with unsigned saturation
'PMULHW Multiply packed signed word integers and store high result
'PMULLW Multiply packed signed word integers and store low result
'PMADDWD Multiply and add packed word integers

'part 4 : MMX Comparison Instructions
'PCMPEQB Compare packed bytes for equal
'PCMPEQW Compare packed words for equal
'PCMPEQD Compare packed doublewords for equal
'PCMPGTB Compare packed signed byte integers for greater than
'PCMPGTW Compare packed signed word integers for greater than
'PCMPGTD Compare packed signed doubleword integers for greater than

'part 5 : MMX Logical Instructions
'PAND Bitwise logical AND
'PANDN Bitwise logical AND NOT
'POR Bitwise logical OR
'PXOR Bitwise logical exclusive OR

'part 6 : MMX Shift and Rotate Instructions
'PSLLW Shift packed words left logical
'PSLLD Shift packed doublewords left logical
'PSLLQ Shift packed quadword left logical
'PSRLW Shift packed words right logical
'PSRLD Shift packed doublewords right logical
'PSRLQ Shift packed quadword right logical
'PSRAW Shift packed words right arithmetic
'PSRAD Shift packed doublewords right arithmetic

'part 7 : MMX State Management Instructions
'EMMS Empty MMX state


End Sub
Private Sub Instr_SSE()
'instructions à partir du Pentium 3 : gestion des floating points simple précision (Single) par paquets
'SSE = Streaming SIMD Extension. Le SSE introduit le registre XMM sur 128 bits


'part 1 : Instruction SSE pour simple précision
'je suppose que les routines DirectX 7 usent et abusent du SSE vu que ca bosse en Single à fond :)

'part 1.1 : SSE Data Transfer Instructions
'SSE data transfer instructions move packed and scalar single-precision floating-point operands between XMM registers and between XMM registers and memory.
'MOVAPS Move four aligned packed single-precision floating-point values between XMM registers or between and XMM register and memory
'MOVUPS Move four unaligned packed single-precision floating-point values between XMM registers or between and XMM register and memory
'MOVHPS Move two packed single-precision floating-point values to an from the high quadword of an XMM register and memory
'MOVHLPS Move two packed single-precision floating-point values from the high quadword of an XMM register to the low quadword of another XMM register
'MOVLPS Move two packed single-precision floating-point values to an from the low quadword of an XMM register and memory
'MOVLHPS Move two packed single-precision floating-point values from the low quadword of an XMM register to the high quadword of another XMM register
'MOVMSKPS Extract sign mask from four packed single-precision floating-point values
'MOVSS Move scalar single-precision floating-point value between XMM registers or between an XMM register and memory

'part 1.2 : SSE Packed Arithmetic Instructions
'ADDPS Add packed single-precision floating-point values
'ADDSS Add scalar single-precision floating-point values
'SUBPS Subtract packed single-precision floating-point values
'SUBSS Subtract scalar single-precision floating-point values
'MULPS Multiply packed single-precision floating-point values
'MULSS Multiply scalar single-precision floating-point values
'DIVPS Divide packed single-precision floating-point values
'DIVSS Divide scalar single-precision floating-point values
'RCPPS Compute reciprocals of packed single-precision floating-point values
'RCPSS Compute reciprocal of scalar single-precision floating-point values
'SQRTPS Compute square roots of packed single-precision floating-point values
'SQRTSS Compute square root of scalar single-precision floating-point values
'RSQRTPS Compute reciprocals of square roots of packed single-precision floating-point Values
'RSQRTSS Compute reciprocal of square root of scalar single-precision floating-point Values
'MAXPS Return maximum packed single-precision floating-point values
'MAXSS Return maximum scalar single-precision floating-point values
'MINPS Return minimum packed single-precision floating-point values
'MINSS Return minimum scalar single-precision floating-point values

'part 1.3 : SSE Comparison Instructions
'CMPPS Compare packed single-precision floating-point values
'CMPSS Compare scalar single-precision floating-point values
'COMISS Perform ordered comparison of scalar single-precision floating-point values and set flags in EFLAGS register
'UCOMISS Perform unordered comparison of scalar single-precision floating-point values and set flags in EFLAGS register

'part 1.4 : SSE Logical Instructions
'ANDPS Perform bitwise logical AND of packed single-precision floating-point values
'ANDNPS Perform bitwise logical AND NOT of packed single-precision floating-point values
'ORPS Perform bitwise logical OR of packed single-precision floating-point values
'XORPS Perform bitwise logical XOR of packed single-precision floating-point values

'part 1.5 : SSE Shuffle and Unpack Instructions
'SHUFPS Shuffles values in packed single-precision floating-point operands
'UNPCKHPS Unpacks and interleaves the two high-order values from two single-precision floating-point operands
'UNPCKLPS Unpacks and interleaves the two low-order values from two single-precision floating-point operands

'part 1.6 : SSE Conversion Instructions
'CVTPI2PS Convert packed doubleword integers to packed single-precision floating-point values
'CVTSI2SS Convert doubleword integer to scalar single-precision floating-point value
'CVTPS2PI Convert packed single-precision floating-point values to packed doubleword integers
'CVTTPS2PI Convert with truncation packed single-precision floating-point values to packed doubleword integers
'CVTSS2SI Convert a scalar single-precision floating-point value to a doubleword integer
'CVTTSS2SI Convert with truncation a scalar single-precision floating-point value to a scalar doubleword integer


'part 2 : SSE MXCSR State Management Instructions (control and status register)
'LDMXCSR Load MXCSR register
'STMXCSR Save MXCSR register state


'part 3 : SSE 64-Bit SIMD Integer Instructions (ajout/supplément aux instructions MMX)
'PAVGB Compute average of packed unsigned byte integers
'PAVGW Compute average of packed unsigned byte integers
'PEXTRW Extract word
'PINSRW Insert word
'PMAXUB Maximum of packed unsigned byte integers
'PMAXSW Maximum of packed signed word integers
'PMINUB Minimum of packed unsigned byte integers
'PMINSW Minimum of packed signed word integers
'PMOVMSKB Move byte mask
'PMULHUW Multiply packed unsigned integers and store high result
'PSADBW Compute sum of absolute differences
'PSHUFW Shuffle packed integer word in MMX register


'part 4 : SSE Cacheability Control, Prefetch, and Instruction Ordering Instructions
'""The cacheability control instructions provide control over the caching of non-temporal data
'  when storing data from the MMX and XMM registers to memory. The PREFETCHh allows data
'  to be prefetched to a selected cache level. The SFENCE instruction controls instruction ordering
'  on store operations.""
'MASKMOVQ Non-temporal store of selected bytes from an MMX register into memory
'MOVNTQ Non-temporal store of quadword from an MMX register into memory
'MOVNTPS Non-temporal store of four packed single-precision floating-point values from an XMM register into memory
'PREFETCHh Load 32 or more of bytes from memory to a selected level of the processor’s cache hierarchy
'SFENCE Serializes store operations


End Sub
Private Sub Instr_SSE2()
'instructions Pentium 4 : gestion des floating points double précision (Double) par paquets, et un peu d'autres trucs

'part 1 : SSE2 Packed and Scalar Double-Precision Floating-Point instructions

'part 1.1 : SSE2 Data Movement Instructions
'MOVAPD Move two aligned packed double-precision floating-point values between XMM registers or between and XMM register and memory
'MOVUPD Move two unaligned packed double-precision floating-point values between XMM registers or between and XMM register and memory
'MOVHPD Move high packed double-precision floating-point value to an from the high quadword of an XMM register and memory
'MOVLPD Move low packed single-precision floating-point value to an from the low quadword of an XMM register and memory
'MOVMSKPD Extract sign mask from two packed double-precision floating-point values
'MOVSD Move scalar double-precision floating-point value between XMM registers or between an XMM register and memory

'part 1.2 : SSE2 Packed Arithmetic Instructions
'ADDPD Add packed double-precision floating-point values
'ADDSD Add scalar double precision floating-point values
'SUBPD Subtract scalar double-precision floating-point values
'SUBSD Subtract scalar double-precision floating-point values
'MULPD Multiply packed double-precision floating-point values
'MULSD Multiply scalar double-precision floating-point values
'DIVPD Divide packed double-precision floating-point values
'DIVSD Divide scalar double-precision floating-point values
'SQRTPD Compute packed square roots of packed double-precision floating-point values
'SQRTSD Compute scalar square root of scalar double-precision floating-point values
'MAXPD Return maximum packed double-precision floating-point values
'MAXSD Return maximum scalar double-precision floating-point values
'MINPD Return minimum packed double-precision floating-point values
'MINSD Return minimum scalar double-precision floating-point values

'part 1.3 : SSE2 Logical Instructions
'ANDPD Perform bitwise logical AND of packed double-precision floating-point values
'ANDNPD Perform bitwise logical AND NOT of packed double-precision floatingpoint values
'ORPD Perform bitwise logical OR of packed double-precision floating-point values
'XORPD Perform bitwise logical XOR of packed double-precision floating-point values

'part 1.4 : SSE2 Compare Instructions
'CMPPD Compare packed double-precision floating-point values
'CMPSD Compare scalar double-precision floating-point values
'COMISD Perform ordered comparison of scalar double-precision floating-point values and set flags in EFLAGS register
'UCOMISD Perform unordered comparison of scalar double-precision floating-point values and set flags in EFLAGS register.

'part 1.5 : SSE2 Shuffle and Unpack Instructions
'SHUFPD Shuffles values in packed double-precision floating-point operands
'UNPCKHPD Unpacks and interleaves the high values from two packed double-precision floating-point operands
'UNPCKLPD Unpacks and interleaves the low values from two packed double-precision floating-point operands

'part 1.6 : SSE2 Conversion Instructions
'CVTPD2PI Convert packed double-precision floating-point values to packed doubleword integers.
'CVTTPD2PI Convert with truncation packed double-precision floating-point values to packed doubleword integers
'CVTPI2PD Convert packed doubleword integers to packed double-precision floating-point values
'CVTPD2DQ Convert packed double-precision floating-point values to packed doubleword integers
'CVTTPD2DQ Convert with truncation packed double-precision floating-point values to packed doubleword integers
'CVTDQ2PD Convert packed doubleword integers to packed double-precision floating-point values
'CVTPS2PD Convert packed single-precision floating-point values to packed double-precision floating-point values
'CVTPD2PS Convert packed double-precision floating-point values to packed single-precision floating-point values
'CVTSS2SD Convert scalar single-precision floating-point values to scalar double-precision floating-point values
'CVTSD2SS Convert scalar double-precision floating-point values to scalar single-precision floating-point values
'CVTSD2SI Convert scalar double-precision floating-point values to a doubleword integer
'CVTTSD2SI Convert with truncation scalar double-precision floating-point values to scalar doubleword integers
'CVTSI2SD Convert doubleword integer to scalar double-precision floating-point value


'part 2 : SSE2 Packed Single-Precision Floating-Point Instructions (ajout, amélioration du SSE)
'CVTDQ2PS Convert packed doubleword integers to packed single-precision floating-point values
'CVTPS2DQ Convert packed single-precision floating-point values to packed doubleword integers
'CVTTPS2DQ Convert with truncation packed single-precision floating-point values to packed doubleword integers


'part 3 : SSE2 128-Bit SIMD Integer Instructions
'MOVDQA Move aligned double quadword.
'MOVDQU Move unaligned double quadword
'MOVQ2DQ Move quadword integer from MMX to XMM registers
'MOVDQ2Q Move quadword integer from XMM to MMX registers
'PMULUDQ Multiply packed unsigned doubleword integers
'PADDQ Add packed quadword integers
'PSUBQ Subtract packed quadword integers
'PSHUFLW Shuffle packed low words
'PSHUFHW Shuffle packed high words
'PSHUFD Shuffle packed doublewords
'PSLLDQ Shift double quadword left logical
'PSRLDQ Shift double quadword right logical
'PUNPCKHQDQ Unpack high quadwords
'PUNPCKLQDQ Unpack low quadwords

'part 4 : SSE2 Cacheability Control and Ordering Instructions
'""SSE2 cacheability control instructions provide additional operations for caching of nontemporal
'  data when storing data from XMM registers to memory. LFENCE and MFENCE
'  provide additional control of instruction ordering on store operations.""
'CLFLUSH Flushes and invalidates a memory operand and its associated cache line from all levels of the processor’s cache hierarchy
'LFENCE Serializes load operations
'MFENCE Serializes load and store operations
'PAUSE Improves the performance of “spin-wait loops”
'MASKMOVDQU Non-temporal store of selected bytes from an XMM register into memory
'MOVNTPD Non-temporal store of two packed double-precision floating-point values from an XMM register into memory
'MOVNTDQ Non-temporal store of double quadword from an XMM register into memory
'MOVNTI Non-temporal store of a doubleword from a general-purpose register into memory


End Sub
Private Sub Instr_SSE3()
'instructions Pentium 4 HT : quelques instructions supplémentaires, pour les autres SIMD, qui manquaient pour aller vite

'part 1 : One x87FPU instruction used in integer conversion
'FISTTP Behaves like the FISTP instruction but uses truncation, irrespective of the rounding mode specified in the floating-point control word (FCW)

'part 2 : One SIMD integer instruction that addresses unaligned data loads
'LDDQU Special 128-bit unaligned load designed to avoid cache line splits

'part 3 : Two SIMD floating-point packed ADD/SUB instructions
'ADDSUBPS Performs single-precision addition on the second and fourth pairs of 32-bit data elements within the operands; single-precision subtraction on the first and third pairs
'ADDSUBPD Performs double-precision addition on the second pair of quadwords, and double-precision subtraction on the first pair

'part 4 :  Four SIMD floating-point horizontal ADD/SUB instructions
'HADDPS Performs a single-precision addition on contiguous data elements. The
'  first data element of the result is obtained by adding the first and second
'  elements of the first operand; the second element by adding the third and
'  fourth elements of the first operand; the third by adding the first and
'  second elements of the second operand; and the fourth by adding the third
'  and fourth elements of the second operand.
'HSUBPS Performs a single-precision subtraction on contiguous data elements. The
'  first data element of the result is obtained by subtracting the second
'  element of the first operand from the first element of the first operand; the
'  second element by subtracting the fourth element of the first operand from
'  the third element of the first operand; the third by subtracting the second
'  element of the second operand from the first element of the second
'  operand; and the fourth by subtracting the fourth element of the second
'  operand from the third element of the second operand.
'HADDPD Performs a double-precision addition on contiguous data elements. The
'  first data element of the result is obtained by adding the first and second
'  elements of the first operand; the second element by adding the first and
'  second elements of the second operand.
'HSUBPD Performs a double-precision subtraction on contiguous data elements. The
'  first data element of the result is obtained by subtracting the second
'  element of the first operand from the first element of the first operand; the
'  second element by subtracting the second element of the second operand
'  from the first element of the second operand.

'part 5 : Three SIMD floating-point LOAD/MOVE/DUPLICATE instructions
'MOVSHDUP Loads/moves 128 bits; duplicating the second and fourth 32-bit data elements
'MOVSLDUP Loads/moves 128 bits; duplicating the first and third 32-bit data elements
'MOVDDUP Loads/moves 64 bits (bits[63:0] if the source is a register) and returns the same 64 bits in both the lower and upper halves of the 128-bit result register; duplicates the 64 bits from the source

'part 6 : Two thread synchronization instructions
'MONITOR Sets up an address range used to monitor write-back stores
'MWAIT Enables a logical processor to enter into an optimized state while waiting for a write-back store to the address range set up by the MONITOR



End Sub
Private Sub Instr_IA32e()
'instructions IA 32 extended : sur 64 bits (Quadwords) avec une pile et des registres approprié.
'c'est les instructions pour le EM64T ...
'il est marrant de voir que Intel ne donne pas beaucoup d'infos dessus dans sa doc officielle :)

'CDQE Convert doubleword to quadword
'CMPSQ Compare string operands
'CMPXCHG16B Compare RDX:RAX with m128
'LODSQ Load qword at address (R)SI into RAX
'MOVSQ Move qword from address (R)SI to (R)DI
'MOVZX (64-bits) Move doubleword to quadword, zero-extension
'STOSQ Store RAX at address RDI
'SWAPGS Exchanges current GS base register value with value in MSR address C0000102H
'SYSCALL Fast call to privilege level 0 system procedures
'SYSRET Return from fast system call


End Sub


Private Function AddInstruct(start As Byte, prefix As Boolean, useint As Boolean, start2 As Byte, direct As Boolean, _
                        usemodrm As Boolean, rmtype As Byte, useregex As Boolean, regexval As Byte) As Long

    tbl_ASMlen = tbl_ASMlen + 1
    AddInstruct = tbl_ASMlen
    With tbl_ASM(tbl_ASMlen)
        .aStart = start
        .bPrefix = prefix
        .cUseInt = useint
        .dStart2 = start2
        .eDirect = direct
        .fUseModRM = usemodrm
        .gRMtype = rmtype
        .hUseRegEx = useregex
        .iRegExVal = regexval
    End With
    
End Function



Sub Decode_ModRM(ByVal ModRM As Byte, ByRef bMod As Byte, ByRef bReg1 As Byte, ByRef bReg2 As Byte)
'décode le byte ModR/M
    bMod = (ModRM / 64) And 3  'nota : si bMod = 2, alors bReg2 est un registre.
    bReg1 = (ModRM / 8) And 7  'aka OpCode - c'est aussi le descripteur /digit
    bReg2 = ModRM And 7        'aka R/M
End Sub

Sub Decode_SIB(ByVal SIB As Byte, ByRef bScale As Byte, ByRef bIndex As Byte, ByRef bBase As Byte)
'décode le byte SIB
    bScale = SIB / 64
    bIndex = (SIB / 8) And 7
    bBase = SIB And 7
End Sub

Function Disassemble(Instruct() As Byte, ByVal iRva As Long, ByRef eRva As Long, ByVal ImageBase As Long, ByRef oCall As Long) As String
'désassemble une instruction, renvoi le code de l'instruction (/!\ sans hex dump ni offset!)
' iRva  : offset de l'instruction (in)
' eRva  : offset de la prochaine instruction (out)
' ImageBase : 00400000h (in) ou autre
' oCall : pour l'interprêteur : l'instruction est un appel vers l'adresse rva oCall
Dim i As Long, j As Long
Static sPrefix As String, PrefixVal As Long
Dim rmMod As Byte, rmReg1 As Byte, rmReg2 As Byte
Dim sibScale As Byte, sibIndex As Byte, sibBase As Byte
    
'IDENTIFICATION
    'cherche l'instruction en fonction du premier octet
    For i = 1 To tbl_ASMlen
        If tbl_ASM(i).aStart = Instruct(1) Then
        
            'est-ce une instruction sur 2 octets ?
            If tbl_ASM(i).cUseInt = False Then
                'non
                
                'est-ce une instruction directe ?
                If tbl_ASM(i).eDirect Then
                    'oui, instruction trouvé et identifié.
                    Disassemble = sPrefix & tbl_ASM(i).zInstruct
                    eRva = iRva + 1
                    sPrefix = ""
                    Exit Function
                
                'est-ce un prefix ?
                If tbl_ASM(i).bPrefix = True Then
                    sPrefix = tbl_ASM(i).zInstruct
                    eRva = iRva + 1
                    Exit Function
                Else
                    sPrefix = ""
                End If
                    
                'est-ce une instruction précisé par Reg1 de l'octet ModR/M ?
                ElseIf tbl_ASM(i).hUseRegEx = True Then
                    Call Decode_ModRM(Instruct(2), rmMod, rmReg1, rmReg2)
                    'est-ce l'instruction identifié par Reg1 ?
                    If tbl_ASM(i).iRegExVal = rmReg1 Then
                        j = 2
                        Exit For 'instruction trouvée
                    End If
                End If
                
            Else
                If tbl_ASM(i).dStart2 = Instruct(2) Then
                
                    'est-ce une instruction directe ?
                    If tbl_ASM(i).eDirect Then
                        'oui, instruction trouvé et identifié.
                        Disassemble = sPrefix & tbl_ASM(i).zInstruct
                        eRva = iRva + 2
                        sPrefix = ""
                        Exit Function
                        
                    'est-ce une instruction précisé par Reg1 de l'octet ModR/M ?
                    ElseIf tbl_ASM(i).hUseRegEx = True Then
                        Call Decode_ModRM(Instruct(3), rmMod, rmReg1, rmReg2)
                        'est-ce l'instruction identifié par Reg1 ?
                        If tbl_ASM(i).iRegExVal = rmReg1 Then
                            j = 3
                            Exit For 'instruction trouvée
                        End If
                    End If
                    
                End If
            
            End If
        End If
    Next i
    
    'gestion de l'instruction
    
    If tbl_ASM(i).fUseModRM = True Then
        Call Decode_ModRM(Instruct(j), rmMod, rmReg1, rmReg2)
        If rmMod = 3 Then
            'addresse directe, 32bit
            
        End If
                
            
        If rmMod = 2 And rmReg2 = 4 Then
            'utilise byte SIB
            Call Decode_SIB(Instruct(j + 1), sibScale, sibIndex, sibBase)
        
        End If
    End If
    
End Function
