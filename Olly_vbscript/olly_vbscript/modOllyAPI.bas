Attribute VB_Name = "modOllyAPI"
Option Explicit
'dzzie@yahoo.com
'http://sandsprite.com

Global frmInstance As New frmOllyScript
Global Bpx_Handler As Long
Global cn As New Connection
Global BpxHandler_Warning As Boolean

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Byte, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Declare Sub SetSyncFlag Lib "ollysync_vbscript.dll" (ByVal action As Long)
Declare Sub DisableBpxHook Lib "ollysync_vbscript.dll" ()
Declare Sub EnableBpxHook Lib "ollysync_vbscript.dll" ()
Declare Sub RefereshDisasm Lib "ollysync_vbscript.dll" ()
Declare Function CurThreadID Lib "ollysync_vbscript.dll" () As Long
Declare Function Follow Lib "ollysync_vbscript.dll" (addr As Long) As Long
Declare Function GetStat Lib "ollysync_vbscript.dll" () As statmode
Declare Function GetRegister Lib "ollysync_vbscript.dll" (ByVal index As Long) As Long
Declare Sub SetRegister Lib "ollysync_vbscript.dll" (ByVal index As regIndex, ByVal val As Long)
Declare Sub RedrawCpu Lib "ollysync_vbscript.dll" ()
Declare Function aGetEIP Lib "ollysync_vbscript.dll" Alias "GetEIP" () As Long
Declare Sub setEIP Lib "ollysync_vbscript.dll" Alias "SetEIP" (ByVal addr As Long)
Declare Function ConfigValue Lib "ollysync_vbscript.dll" (ByVal var As configVals) As Long
Declare Function OpenEXE Lib "ollysync_vbscript.dll" (ByVal fpath As String) As Boolean
Declare Function DecodeAsc Lib "ollysync_vbscript.dll" (offset As Long, ByVal buf As String, buflen As Long) As Long
Declare Function DecodeUni Lib "ollysync_vbscript.dll" (offset As Long, ByVal buf As String, buflen As Long) As Long
Declare Function FindImportName Lib "ollysync_vbscript.dll" (ByVal buf As String, startAddr As Long, endAddr As Long) As Long

Declare Sub DelHwrBPXAddr Lib "ollysync_vbscript.dll" (offset As Long)
Declare Sub DelHwrBPXIndex Lib "ollysync_vbscript.dll" (index As Long)

'int __stdcall GetModuleBaseAddress(void){
Public Declare Function aGetModuleBaseAddress _
                 Lib "ollysync_vbscript.dll" _
                 Alias "GetModuleBaseAddress" () _
                 As Long

'int __stdcall GetModuleSize(void){
Declare Function aGetModuleSize _
                Lib "ollysync_vbscript.dll" _
                Alias "GetModuleSize" () _
                As Long

'int __stdcall GetName(int *offset, int *type, LPSTR pszString){
Declare Function aGetName _
                 Lib "ollysync_vbscript.dll" _
                 Alias "GetName" _
                 (offset As Long, typ As nmNames, ByVal buf As String) _
                 As Long

'int __stdcall SetName(int *offset, int *type, LPSTR pszString){
Declare Function aSetName _
                Lib "ollysync_vbscript.dll" _
                Alias "SetName" _
                (offset As Long, typ As nmNames, ByVal buf As String) _
                As Long

'int __stdcall ProcBegin(int *addr){
Declare Function aProcBegin _
                 Lib "ollysync_vbscript.dll" _
                 Alias "ProcBegin" (addr As Long) As Long

'int __stdcall ProcEnd(int *addr){
Declare Function aProcEnd _
                 Lib "ollysync_vbscript.dll" _
                 Alias "ProcEnd" (addr As Long) As Long

'int __stdcall PrevProc(int *addr){
Declare Function aPrevProc _
                 Lib "ollysync_vbscript.dll" _
                 Alias "PrevProc" (addr As Long) As Long

'int __stdcall NextProc(int *addr){
Declare Function aNextProc _
                 Lib "ollysync_vbscript.dll" _
                 Alias "NextProc" (addr As Long) As Long

'void __stdcall DoAnimate(int *animation){
Declare Sub aDoAnimate _
            Lib "ollysync_vbscript.dll" _
            Alias "DoAnimate" (animation As anmNames)

'int __stdcall SetBpx(int *addr, int *type, char *cmd){
Declare Function aSetBpx _
                 Lib "ollysync_vbscript.dll" _
                 Alias "SetBpx" _
                 (addr As Long, typ As Long, cmd As Byte) As Long

'int __stdcall SetHdwBpx(int *addr, int *size, int *type){
Declare Function aSetHdwBpx _
                 Lib "ollysync_vbscript.dll" _
                 Alias "SetHdwBpx" _
                 (addr As Long, size As Long, typ As Long) As Long

'int __stdcall GetCmd(int *ip, char *cmd){
Declare Function aGetByteCode _
                 Lib "ollysync_vbscript.dll" _
                 Alias "GetByteCode" _
                 (ip As Long, ByVal buf As String) As Long

'int __stdcall ReadMem(char *buf, int*addr, int *size, int *mode){
Declare Function aReadMem _
                 Lib "ollysync_vbscript.dll" _
                 Alias "ReadMem" _
                 (buf As Byte, addr As Long, size As Long, mode As Long) _
                 As Long

'int __stdcall WriteMem(char *buf, int*addr, int *size, int *mode){
Declare Function aWriteMem _
                 Lib "ollysync_vbscript.dll" _
                 Alias "WriteMem" _
                 (buf As Byte, addr As Long, size As Long, mode As Long) _
                 As Long

'int __stdcall DoGo(int *threadid, int *tilladdr, int *stepmode, int *givechance, int *backupargs){
Declare Function DoGo Lib "ollysync_vbscript.dll" _
               (threadId As Long, tiladdr As Long, stepMode As Long, givechance As Long, backuparg As Long) _
                As Long
                
'int __stdcall GetAsm(char *src, int *srcsize, int *srcip, LPSTR pszString){
Declare Function aGetAsm Lib "ollysync_vbscript.dll" _
                    Alias "GetAsm" _
                    (buf As Byte, bufSize As Long, srcOffset As Long, ByVal retbuf As String) As Long
                
'int __stdcall Asm(LPSTR cmd, int *addr, LPSTR errBuffer){
Declare Function Asm Lib "ollysync_vbscript.dll" _
                   (ByVal asmTxt As String, offset, ByVal errBuf As String, ByVal codeBuf As String) As Long
                
'int __stdcall DecodeAddr(int *addr, int *base, int*addrmode, char*buf, int *buflen, char *combuf){
Declare Function DecodeAddr Lib "ollysync_vbscript.dll" _
                   (adr As Long, base As Long, adrMode As adrModes, _
                     ByVal buf As String, buflen As Long, ByVal commentBuffer As String) As Long
                   
'int __stdcall ImageBaseAndNameForEA(int ea, int *outValBase, char* buf, int bufSize){
Declare Function ImageBaseAndNameForEA Lib "ollysync_vbscript.dll" _
                   (ByVal adr As Long, ByRef outValBase As Long, _
                     ByVal buf As String, ByVal buflen As Long) As Long

           
Public Const MM_RESTORE = 1   ' Restore or remove INT3 breakpoints
Public Const MM_SILENT = 2    ' Don't display error message
Public Const MM_DELANAL = 4   ' Delete analysis from the memory
           
Enum regIndex
    REG_EAX = 0
    REG_ECX = 1
    REG_EDX = 2
    REG_EBX = 3
    REG_Esp = 4
    REG_Ebp = 5
    REG_Esi = 6
    REG_Edi = 7
End Enum

Enum configVals
'        VAL_HINST = 1          ' Current program instance
        VAL_HWMAIN = 2         ' Handle of the main window
'        VAL_HWCLIENT = 3       ' Handle of the MDI client window
'        VAL_NCOLORS = 4        ' Number of common colors
'        VAL_COLORS = 5         ' RGB values of common colors
'        VAL_BRUSHES = 6        ' Handles of common color brushes
'        VAL_PENS = 7           ' Handles of common color pens
'        VAL_NFONTS = 8         ' Number of common fonts
'        VAL_FONTS = 9          ' Handles of common fonts
'        VAL_FONTNAMES = 10     ' Internal font names
'        VAL_FONTWIDTHS = 11    ' Average widths of common fonts
'        VAL_FONTHEIGHTS 12             ' Average heigths of common fonts
'        VAL_NFIXFONTS = 13     ' Actual number of fixed-pitch fonts
'        VAL_DEFFONT = 14       ' Index of default font
'        VAL_NSCHEMES = 15      ' Number of color schemes
'        VAL_SCHEMES = 16       ' Color schemes
'        VAL_DEFSCHEME = 17     ' Index of default colour scheme
'        VAL_DEFHSCROLL = 18    ' Default horizontal scroll
'        VAL_RESTOREWINDOWPOS 19        ' Restore window positions from .ini
        VAL_HPROCESS = 20      ' Handle of Debuggee
        VAL_PROCESSID = 21     ' Process ID of Debuggee
        VAL_HMAINTHREAD = 22            ' Handle of main thread
        VAL_MAINTHREADID = 23           ' Thread ID of main thread
'        VAL_MAINBASE = 24      ' Base of main module in the process
        VAL_PROCESSNAME = 25            ' Name of the active process
        VAL_EXEFILENAME = 26            ' Name of the main debugged file
        VAL_CURRENTDIR = 27    ' Current directory for debugged process
        VAL_SYSTEMDIR = 28     ' Windows system directory
'        VAL_DECODEANYIP 29             ' Decode registers dependless on EIP
'        VAL_PASCALSTRINGS 30           ' Decode Pascal-style string constants
'        VAL_ONLYASCII = 31     ' Only printable ASCII chars in dump
'        VAL_DIACRITICALS 32            ' Allow diacritical symbols in strings
'        VAL_GLOBALSEARCH 33            ' Search from the beginning of block
'        VAL_ALIGNEDSEARCH 34           ' Search aligned to item's size
'        VAL_IGNORECASE = 35    ' Ignore case in string search
'        VAL_SEARCHMARGIN 36            ' Floating search allows error margin
'        VAL_KEEPSELSIZE 37             ' Keep size of hex edit selection
'        VAL_MMXDISPLAY = 38    ' MMX display mode in dialog
'        VAL_WINDOWFONT = 39    ' Use calling window's font in dialog
'        VAL_TABSTOPS = 40      ' Distance between tab stops
'        VAL_MODULES = 41       ' Table of modules (.EXE and .DLL)
'        VAL_MEMORY = 42        ' Table of allocated memory blocks
'        VAL_THREADS = 43       ' Table of active threads
        VAL_BREAKPOINTS = 44            ' Table of active breakpoints
        VAL_REFERENCES = 45    ' Table with found references
'        VAL_SOURCELIST = 46    ' Table of source files
'        VAL_WATCHES = 47       ' Table of watches
'        VAL_CPUFEATURES 50             ' CPU feature bits
'        VAL_TRACEFILE = 51     ' Handle of run trace log file
'        VAL_ALIGNDIALOGS 52            ' Whether to align dialogs
'        VAL_CPUDASM = 53       ' Dump descriptor of CPU Disassembler
'        VAL_CPUDDUMP = 54      ' Dump descriptor of CPU Dump
'        VAL_CPUDSTACK = 55     ' Dump descriptor of CPU Stack
'        VAL_APIHELP = 56       ' Name of selected API help file
        VAL_HARDBP = 57        ' Whether hardware breakpoints enabled
End Enum
 
Enum adrModes
        'ADC_DEFAULT = &H0              ' Default decoding mode
        ADC_DIFFMOD = &H1              ' Show module only if different
        ADC_NOMODNAME = &H2            ' Never show module name
        ADC_VALID = &H4           ' Only decode if allocated memory
        ADC_INMODULE = &H8             ' Only decode if in some module
        ADC_SAMEMOD = &H10             ' Decode only address in same module
        ADC_SYMBOL = &H20         ' Only decode if symbolic name
        ADC_JUMP = &H40           ' Check if points to JMP/CALL command
        ADC_OFFSET = &H80         ' Check if symbol for data
        ADC_STRING = &H100        ' Check if pointer to ASCII or UNICODE
        ADC_ENTRY = &H200          ' Check if entry to subroutine
        'ADC_UPPERCASE = &H400          ' First letter in uppercase if possible
        'ADC_WIDEFORM = &H800           ' Extended form of decoded name
        'ADC_NONTRIVIAL = &H1000        ' Name + non-zero offset
        'ADC_DYNAMIC = &H2000           ' JMP/CALL to dynamically loaded name
End Enum

           
Enum statmode
      STAT_NONE = 0         ' Thread/process is empty
      STAT_STOPPED = 1      ' Thread/process suspended
      STAT_EVENT = 2        ' Processing debug event, process paused
      STAT_RUNNING = 3      ' Thread/process running
      STAT_FINISHED = 4     ' Process finished
      STAT_CLOSING = 5      ' Process is requested to terminate
End Enum
                
Enum stepModes
    STEP_SAME = 0          '         // Same action as on previous call
    STEP_RUN = 1           '         // Run program
    STEP_OVER = 2          '         // Step over
    STEP_IN = 3            '         // Step in
    STEP_SKIP = 4          '         // Skip sequence
End Enum



Enum hwBpTypes
    HB_CODE = 1     'Active on command execution
    HB_ACCESS = 2   'Active on read/write access
    HB_WRITE = 3    'Active on write access
End Enum

Enum nmNames
'       NM_NONAME = &H0                ' Undefined name
       NM_ANYNAME = &HFF              ' Name of any type
' Names saved in the data file of module they appear.
       NM_LABEL = &H31                ' User-defined label
       NM_EXPORT = &H32               ' Exported (global) name
       NM_IMPORT = &H33               ' Imported name
       NM_LIBRARY = &H34              ' Name from library or object file
       NM_CONST = &H35                ' User-defined constant
       NM_COMMENT = &H36              ' User-defined comment
'       NM_LIBCOMM = &H37              ' Comment from library or object file
'       NM_BREAK = &H38                ' Condition related with breakpoint
       NM_ARG = &H39                  ' Arguments decoded by analyzer
       NM_ANALYSE = &H3A              ' Comment added by analyzer
'       NM_BREAKEXPR = &H3B            ' Expression related with breakpoint
'       NM_BREAKEXPL = &H3C            ' Explanation related with breakpoint
'       NM_ASSUME = &H3D               ' Assume function with known arguments
       NM_STRUCT = &H3E               ' Code structure decoded by analyzer
'       NM_CASE = &H3F                 ' Case description decoded by analyzer
' Names saved in the data file of main module.
'       NM_INSPECT = &H40              ' Several last inspect expressions
'       NM_WATCH = &H41                ' Watch expressions
'       NM_ASM = &H42                  ' Several last assembled strings
'       NM_FINDASM = &H43              ' Several last find assembler strings
'       NM_LASTWATCH = &H48            ' Several last watch expressions
'       NM_SOURCE = &H49               ' Several last source search strings
'       NM_REFTXT = &H4A               ' Several last ref text search strings
'       NM_GOTO = &H4B                 ' Several last expressions to follow
'       NM_GOTODUMP = &H4C             ' Several expressions to follow in Dump
'       NM_TRPAUSE = &H4D              ' Several expressions to pause trace
' Pseudonames.
       NM_IMCALL = &HFE               ' Intermodular call

'       NMHISTORY = &H40               ' Converts NM_xxx to type of init list
End Enum


Enum anmNames ' Status of animation or trace.
'       ANIMATE_OFF = 0                 ' No animation
       ANIMATE_IN = 1                  ' Animate into
       ANIMATE_OVER = 2                ' Animate over
       ANIMATE_RET = 3                ' Execute till RET
       ANIMATE_SKPRET = 4              ' Skip RET instruction
       ANIMATE_USER = 5                ' Execute till user code
'       ANIMATE_TRIN = 6                ' Run trace in
'       ANIMATE_TROVER = 7              ' Run trace over
'       ANIMATE_STOP = 8                ' Gracefully stop animation
End Enum

Enum bpTypes
'        TY_STOPAN =&H80               ' Stop animation if TY_ONESHOT
'        TY_SET =&H100                 ' Code INT3 is in memory
        TY_ACTIVE = &H200             ' Permanent breakpoint
        TY_DISABLED = &H400           ' Permanent disabled breakpoint
        TY_ONESHOT = &H800            ' Temporary stop
        TY_TEMP = &H1000              ' Temporary breakpoint
'        TY_KEEPCODE =&H2000           ' Set and keep command code
'        TY_KEEPCOND =&H4000           ' Keep condition unchanged (0: remove)
'        TY_NOUPDATE =&H8000           ' Don't redraw breakpoint window
'        TY_RTRACE =&H10000            ' Pseudotype of run trace breakpoint

'dzzie@yahoo.com
'http://sandsprite.com
End Enum


Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    Dim X As Long
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
init:     ReDim ary(0): ary(0) = Value
End Sub


Function GenericGet(offset As Long, typ As nmNames) As String
    Dim retLen As Long
    Dim buf As String
    
    buf = String(257, Chr(0))
    retLen = aGetName(offset, typ, buf)

    If retLen > 0 Then
        buf = Mid(buf, 1, retLen)
    Else
        buf = Empty
    End If

    GenericGet = buf

End Function

Function GenericSet(offset As Long, typ As nmNames, buf As String) As Long
    GenericSet = aSetName(offset, typ, buf)
End Function

Function GenericBPX(offset As Long, typ As bpTypes) As Long
      GenericBPX = aSetBpx(offset, typ, 0)
End Function

Function GenericHwBPX(offset As Long, typ As hwBpTypes) As Long
     GenericHwBPX = aSetHdwBpx(offset, 1, typ)
End Function

Function GenericGo(stepMode As stepModes, Optional threadId As Long = 0)
    GenericGo = DoGo(threadId, 0, stepMode, 0, 0)
End Function

Function GenericReg(reg As regIndex) As Long
    GenericReg = GetRegister(CLng(reg))
End Function

Function GenericDecode(mode As adrModes, adr As Long, commentBuffer As String, Optional base As Long = 0, Optional stringBuf As Long = 255) As String
    
    Dim buf As String, l As Long
    buf = String(stringBuf, Chr(0))
    commentBuffer = String(255, Chr(0))
    
    l = DecodeAddr(adr, base, adr, buf, stringBuf, commentBuffer)
    
    'MsgBox l & " " & buf & vbCrLf & _
           "Comment:" & commentBuffer
    
    If l > 0 Then
        buf = Mid(buf, 1, l)
    End If
    
    GenericDecode = buf
    
End Function
 
Function GetImageBaseAndNameForVA(ea, ByRef outBase As Long, ByRef outString As String) As Boolean
    Dim ret As Long
    
    outString = String(255, " ")
    
    ret = ImageBaseAndNameForEA(ea, outBase, outString, 255)
    GetImageBaseAndNameForVA = IIf(ret = 1, True, False)
    outString = Trim(Replace(outString, Chr(0), ""))
    
End Function
