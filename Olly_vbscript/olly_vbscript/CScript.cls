VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COlly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
'dzzie@yahoo.com
'http://sandsprite.com


 Sub ShowBpxHandlerWarning(Optional enabled = True)
    BpxHandler_Warning = CBool(enabled)
 End Sub
 
 Sub Nop(offset, length)
    On Error Resume Next
    Dim b() As Byte
    b() = StrConv(String(length, Chr(&H90)), vbFromUnicode)
    WriteMem b, offset
    Refresh
 End Sub
 
 Sub ClearLog()
    frmInstance.txtLog = Empty
 End Sub
 
 Sub LogIt(X As Variant)
        frmInstance.txtLog.SelStart = Len(frmInstance.txtLog.Text)
        frmInstance.txtLog.SelText = CStr(X) & vbCrLf
 End Sub
 
 Sub l(X)
    LogIt X
 End Sub
 
 Sub EnableBpxHook()
    modOllyAPI.EnableBpxHook
 End Sub

Sub DisableBpxHook()
    modOllyAPI.DisableBpxHook
End Sub

Sub ShowUI()
    frmInstance.Show 1
End Sub

Function ModBaseAddr() As Long
    ModBaseAddr = aGetModuleBaseAddress()
End Function

Function ModSize() As Long
    ModSize = aGetModuleSize()
End Function

Function ProcBegin(addr) As Long
    ProcBegin = aProcBegin(CLng(addr))
End Function

Function ProcEnd(addr) As Long
    ProcEnd = aProcEnd(CLng(addr))
End Function

Function PrevProc(addr) As Long
    PrevProc = aPrevProc(CLng(addr))
End Function

Function NextProc(addr) As Long
     NextProc = aNextProc(CLng(addr))
End Function


Sub StepInto()
     'aDoAnimate ANIMATE_IN
     GenericGo STEP_IN
     frmInstance.Visible = False
End Sub
       
Sub StepOver()
     'aDoAnimate ANIMATE_OVER
     GenericGo STEP_OVER
     frmInstance.Visible = False
End Sub
       
Sub ExecTillRet()
     aDoAnimate ANIMATE_RET
     frmInstance.Visible = False
End Sub
       
Sub ExecTillUserCode()
    aDoAnimate ANIMATE_USER
    frmInstance.Visible = False
End Sub

Sub Go()

    GenericGo STEP_RUN
    frmInstance.Visible = False
    
End Sub




'--------------------------------------------------------
Function GetLabel(offset) As String
    On Error Resume Next
    GetLabel = GenericGet(CLng(offset), NM_LABEL)
End Function

Function GetComment(offset) As String
    On Error Resume Next
    GetComment = GenericGet(CLng(offset), NM_COMMENT)
End Function

Function GetAny(offset) As String
    On Error Resume Next
    GetAny = GenericGet(CLng(offset), NM_ANYNAME)
End Function

Function SetLabel(offset, val)
    SetLabel = GenericSet(CLng(offset), NM_LABEL, CStr(val))
    Refresh
End Function

Function SetComment(offset, val)
    SetComment = GenericSet(CLng(offset), NM_COMMENT, CStr(val))
    Refresh
End Function
 

'--------------------------------------------------------------
Function SetBpx(adr) As Long
    On Error Resume Next
    GenericBPX CLng(adr), TY_ACTIVE
    Refresh
End Function

Function RemoveBpx(adr) As Long
    On Error Resume Next
    GenericBPX CLng(adr), TY_DISABLED
    Refresh
End Function

Function OneShotBpx(adr) As Long
    On Error Resume Next
    GenericBPX CLng(adr), TY_ONESHOT
    Refresh
End Function


Function SetHdwBpAccess(adr) As Long
    On Error Resume Next
    SetHdwBpAccess = GenericHwBPX(CLng(adr), HB_ACCESS)
    Refresh
End Function

Function SetHdwBpWrite(adr) As Long
    On Error Resume Next
    SetHdwBpWrite = GenericHwBPX(CLng(adr), HB_WRITE)
    Refresh
End Function

Function SetHdwBpExecution(adr) As Long
    On Error Resume Next
    SetHdwBpExecution = GenericHwBPX(CLng(adr), HB_CODE)
    Refresh
End Function

Function GetByteCode(offset, Optional strAsm) 'As Byte() (0 based)
    Dim buf As String, leng As Long, b() As Byte, r As String
    On Error Resume Next
    
    buf = String(17, Chr(0))
    leng = aGetByteCode(CLng(offset), buf)
     
    b() = StrConv(buf, vbFromUnicode)
        
    r = String(257, Chr(0))
        
    leng = aGetAsm(b(0), UBound(b), CLng(offset), r)
    
    strAsm = r
    'MsgBox leng & " " & Hex(b(0)) & " " & r
    
    If leng < 1 Then
        leng = 16
    Else
        leng = leng - 1
    End If
    
    Dim v() As Variant, i As Integer
    ReDim v(leng)
    
    For i = 0 To leng
        v(i) = b(i)
    Next
    
    GetByteCode = v()
        
End Function

Function InstLen(offset) As Long '1 based
    On Error Resume Next
    InstLen = UBound(GetByteCode(offset)) + 1
End Function


Sub Refresh()
    RefereshDisasm
End Sub

Function CurrentThreadID() As Long
    CurrentThreadID = CurThreadID()
End Function

Function FollowCall(addr) As Long
    On Error Resume Next
    FollowCall = Follow(CLng(addr))
    Refresh
End Function

Function ReadMem(addr, size) 'As Byte()
    
    Dim buf As String
    Dim v() As Variant
    Dim b() As Byte
    Dim i As Integer
    
    On Error Resume Next
    
    buf = String(size + 1, Chr(0))
    ReDim v(size)
    ReDim b(size)
    
    Call aReadMem(b(0), CLng(addr), size, MM_RESTORE Or MM_SILENT)
    
    For i = 0 To size
        v(i) = b(i)
    Next
    
    ReadMem = v()
    
End Function

Function WriteMem(buf, addr) As Long
    Dim b() As Byte, i As Integer
    
    On Error Resume Next
    
    ReDim b(UBound(buf))
    
    For i = 0 To UBound(buf)
        b(i) = CByte(buf(i))
    Next
    
    If Err.Number > 0 Then
        MsgBox "WriteMem Bad input! Err:" & Err.Description
        Exit Function
    End If
    
    WriteMem = aWriteMem(b(0), addr, UBound(b) + 1, MM_RESTORE Or MM_SILENT)
    Refresh
    
End Function

Function GetAsm(offset) As String
    Dim r As String
    
    GetByteCode offset, r
    
    If InStr(r, Chr(0)) > 0 Then
        r = Mid(r, 1, (InStr(r, Chr(0)) - 1))
    End If
    
    GetAsm = r
    
End Function

Function Asm(offset, cmd)
    Dim cod As String, errs As String, opBuf As String
    Dim codeleng As Long
    
    cod = CStr(cmd)
    errs = String(257, Chr(0))
    opBuf = String(257, Chr(0))
    
    On Error Resume Next
    codeleng = modOllyAPI.Asm(cod, CLng(offset), errs, opBuf)
    
    If codeleng = 0 Then
        MsgBox "Error: " & errs
        Exit Function
    End If
    
    Dim b() As Byte
    opBuf = Mid(opBuf, 1, codeleng)
    b() = StrConv(opBuf, vbFromUnicode)
    
    WriteMem b, offset
    Asm = codeleng
    Refresh
End Function

Function BlockAsm(ByVal offset, cmds)

    If InStr(cmds, ";") > 0 Then
        cmds = Replace(cmds, ";", vbCrLf)
    End If
    
    Dim tmp() As String
    Dim leng As Long, i As Long, blockSize As Long
    
    tmp = Split(cmds, vbCrLf)
    
    For i = 0 To UBound(tmp)
        leng = Asm(offset + blockSize, tmp(i))
        blockSize = blockSize + leng
    Next
    
End Function

'----------------------------------------------------------------
'Registers stuff
'----------------------------------------------------------------

Function EIP() As Long
    EIP = aGetEIP()
End Function

Sub setEIP(X)
    On Error Resume Next
    modOllyAPI.setEIP CLng(X)
    RedrawCpu
End Sub

 Property Get EAX()
    EAX = GenericReg(REG_EAX)
 End Property
 
Sub SetEAX(X)
    On Error Resume Next
    SetRegister REG_EAX, CLng(X)
    RedrawCpu
End Sub
 
 Property Get ECX() As Long
    ECX = GenericReg(REG_ECX)
 End Property
 
 Sub setECX(X)
    On Error Resume Next
    SetRegister REG_ECX, CLng(X)
    RedrawCpu
 End Sub
 
 Property Get EDX() As Long
    EDX = GenericReg(REG_EDX)
 End Property
 
 Sub setEDX(X)
    On Error Resume Next
    SetRegister REG_EDX, CLng(X)
    RedrawCpu
 End Sub
 
 Property Get EBX() As Long
    EBX = GenericReg(REG_EBX)
 End Property
 
 Sub setEBX(X)
    On Error Resume Next
    SetRegister REG_EBX, CLng(X)
    RedrawCpu
 End Sub
 
 Property Get ESP() As Long
    ESP = GenericReg(REG_Esp)
 End Property
 
 Sub setEsp(X)
    On Error Resume Next
    SetRegister REG_Esp, CLng(X)
    RedrawCpu
 End Sub
 
 Property Get EBP() As Long
    EBP = GenericReg(REG_Ebp)
 End Property
 
 Sub setEBP(X)
    On Error Resume Next
    SetRegister REG_Ebp, CLng(X)
    RedrawCpu
 End Sub
 
 Property Get ESI() As Long
    ESI = GenericReg(REG_Esi)
 End Property
 
 Sub setESI(X)
    On Error Resume Next
    SetRegister REG_Esi, CLng(X)
    RedrawCpu
 End Sub
 
 Property Get EDI() As Long
    EDI = GenericReg(REG_Edi)
 End Property
 
 Sub setEDI(X)
    On Error Resume Next
    SetRegister REG_Edi, CLng(X)
    RedrawCpu
 End Sub
 
 Property Get BpxHandler()
    BpxHandler = Bpx_Handler
 End Property
    
Sub SetBpxHandler(X)
    On Error Resume Next
    BpxHandler = CLng(X)
End Sub

Function Dump(start, length, fpath) As Boolean
    Dim pHandle As Long
    Dim b() As Byte
    Dim f As Long
    
    On Error GoTo hell
    
    If Len(fpath) = 0 Then Exit Function
    
    f = FreeFile
    
    ReDim b(length - 1) '0 based
        
    pHandle = ConfigValue(VAL_HPROCESS)
    ReadProcessMemory pHandle, CLng(start), b(0), CLng(length), CLng(length)
    
    Open fpath For Binary As f
    Put f, , b()
    Close f
    
    Dump = True
    
    Exit Function
hell:
    Dump = False
End Function

Sub Exec(cmd)
   On Error Resume Next
   Shell cmd, vbNormalFocus
End Sub

Function OpenFile(path) As Boolean
    OpenFile = OpenEXE(CStr(path))
End Function

Function DecodeString(offset, Optional commentBuffer) As String
    On Error Resume Next
    DecodeString = GenericDecode(ADC_STRING, CLng(offset), CStr(commentBuffer))
End Function

Sub RefreshCPU()
    RedrawCpu
End Sub

Function DecodeAscii(offset, Optional bufSize = 257) As String
    Dim buf As String, ret As Long
    
    On Error Resume Next
    
    buf = String(bufSize, Chr(0))
    ret = DecodeAsc(CLng(offset), buf, bufSize)
    
    If ret > 1 Then buf = Mid(buf, 1, ret)
    
    DecodeAscii = buf
    
End Function

Function DecodeUniStr(offset, Optional bufSize = 257) As String
    Dim buf As String, ret As Long
    
    On Error Resume Next
    
    buf = String(bufSize, Chr(0))
    ret = DecodeUni(CLng(offset), buf, bufSize)
    
    If ret > 1 Then buf = Mid(buf, 1, ret)
    
    DecodeUniStr = buf
    
End Function

Function FindImportAddress(impName) As Long
    
    Dim start As Long
    start = ModBaseAddr
    
    FindImportAddress = FindImportName(CStr(impName), start, start + ModSize)
    
End Function

Function HideDebugger(Optional breakOnRead = False)
    Const PEB_DBGR As Long = &H7FFDF002
        
    On Error Resume Next
    
    WriteMem Array(0), PEB_DBGR
    If breakOnRead Then SetHdwBpAccess PEB_DBGR

End Function
 
Sub RemoveHwdBpx(valu, Optional byIndex = False)
    On Error Resume Next
    If byIndex Then
        DelHwrBPXIndex CLng(valu)
    Else
        DelHwrBPXAddr CLng(valu)
    End If
End Sub

Sub SetMultiBpx(ParamArray bpxs())
    On Error Resume Next
    Dim i As Integer
    For i = 0 To UBound(bpxs)
        SetBpx CLng(bpxs(i))
    Next
End Sub

Sub RemMultiBpx(ParamArray bpxs())
    On Error Resume Next
    Dim i As Integer
    For i = 0 To UBound(bpxs)
        RemoveBpx CLng(bpxs(i))
    Next
End Sub

'dzzie@yahoo.com
'http://sandsprite.com
