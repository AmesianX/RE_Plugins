VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIDAScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ipc As CIpc
Attribute ipc.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ipc = New CIpc
    ipc.FindClient "IDA_SERVER"
    ipc.Listen Form1.hwnd
End Sub

Property Get isUp() As Boolean
    isUp = IIf(IsWindow(ipc.RemoteHWND) = 1, True, False)
End Property

Function alert(msg)
    If IsArray(msg) Then
        MsgBox "Array: " & vbCrLf & vbCrLf & Join(msg, vbCrLf), vbInformation, "IDAJScript"
    Else
        MsgBox msg, vbInformation, "IDAJScript"
    End If
End Function

Function Message(msg As String)
    ipc.Send "msg:" & msg
End Function

Function MakeStr(va, Optional ascii As Boolean = True)
    ipc.Send "makestring:" & va & ":" & IIf(ascii, "ascii", "unicode")
End Function

Function MakeUnk(va, size)
    ipc.Send "makeunk:" & va & ":" & size
End Function

Property Get LoadedFile() As String
    LoadedFile = ipc.SendCmdRecvText("loadedfile:" & Form1.hwnd)
End Property

Sub t(data As String)
    Form1.Text1.Text = Form1.Text1.Text & vbCrLf & data
    Form1.Text1.SelStart = Len(Form1.Text1.Text)
End Sub

Sub ClearLog()
    Form1.Text1.Text = Empty
End Sub

Function PatchString(va, str, Optional isUnicode = False)
    
    Dim b() As Byte, i As Long
    b() = StrConv(str, vbFromUnicode)
    For i = 0 To UBound(b)
        PatchByte va, b(i)
        If isUnicode Then PatchByte va + 1, 0
        va = IIf(isUnicode, va + 2, va + 1)
    Next
    
End Function

Function PatchByte(va, newVal)
    If CInt(newVal) > 255 Then Err.Raise 1, "PatchByte", "byte value must be 0-254"
    ipc.Send "patchbyte:" & va & ":" & newVal
End Function



Function intToHex(x)
    On Error Resume Next
    intToHex = Hex(x)
    If Len(intToHex) = 1 Then intToHex = "0" & intToHex
    If Err.Number <> 0 Then intToHex = x 'hex can throw overflow bleh...
End Function

Function GetAsm(va)
    GetAsm = ipc.SendCmdRecvText("getasm:" & va & ":" & Form1.hwnd)
End Function

Function InstSize(offset)
    Dim n
    n = NextEA(offset)
    If n = 0 Then Exit Function
    InstSize = n - offset
End Function

'Function InstSize(offset)
'    'there has to be a real way to do this
'    Dim x As String, tmp As String, i As Long, n As String
'    Dim firstea As Long, secondea As Long
'    Dim leng As Long
'
'    leng = 40
'    firstea = 0
'    secondea = 0
'    For i = 0 To leng - 1
'        tmp = GetAsm(offset + i)
'        If tmp = "Fail" Then tmp = Empty
'        If Len(tmp) > 0 Then
'            If firstea = 0 Then
'                firstea = offset + i
'            ElseIf secondea = 0 Then
'                 secondea = offset + i
'            End If
'            If firstea > 0 And secondea > 0 Then Exit For
'        End If
'    Next
'
'    InstSize = secondea - firstea
'
'End Function

Function XRefsTo(offset) 'returns CSV list
    Dim tmp
    tmp = ipc.SendCmdRecvText("refsto:" & offset & ":" & Form1.hwnd)
    If Right(tmp, 1) = "," Then tmp = Mid(tmp, 1, Len(tmp) - 1)
    'tmp = Split(tmp, ",") javascript cant deal with string() must be variant()?
    XRefsTo = tmp
End Function

Function XRefsFrom(offset) 'returns CSV list
    Dim tmp
    tmp = ipc.SendCmdRecvText("refsfrom:" & offset & ":" & Form1.hwnd)
    If Right(tmp, 1) = "," Then tmp = Mid(tmp, 1, Len(tmp) - 1)
    'tmp = Split(tmp, ",") javascript cant deal with string() must be variant()?
    XRefsFrom = tmp
End Function


Function GetName(offset)
    GetName = ipc.SendCmdRecvText("getname:" & offset & ":" & Form1.hwnd)
    If GetName = "NULL" Then GetName = ""
End Function

 Function FunctionName(functionIndex)
    FunctionName = ipc.SendCmdRecvText("funcname:" & functionIndex & ":" & Form1.hwnd)
 End Function

'
Function HideBlock(offset, leng)
    Dim x, i As Long
    x = offset + leng
    For i = offset To x
        HideEA i
    Next
    refresh
End Function

Function ShowBlock(offset, leng)
    Dim x, i As Long
    x = offset + leng
    For i = offset To x
        ShowEA i
    Next
    refresh
End Function

Sub Setname(offset, name)
    ipc.Send "setname:" & offset & ":" & CStr(name)
 End Sub

Sub AddComment(offset, comment)
    ipc.Send "addcomment:" & offset & ":" & comment
End Sub

Function GetComment(offset)
    GetComment = ipc.SendCmdRecvText("getcomment:" & offset & ":" & Form1.hwnd)
    If GetComment = "NULL" Then GetComment = Empty
End Function


 Sub AddCodeXRef(offset, tova)
    ipc.Send "addcodexref:" & offset & ":" & tova
 End Sub
 
  Sub AddDataXRef(offset, tova)
    ipc.Send "adddataxref:" & offset & ":" & tova
 End Sub
 
 Sub DelCodeXRef(offset, tova)
    ipc.Send "delcodexref:" & offset & ":" & tova
 End Sub
 
 Sub DelDataXRef(offset, tova)
    ipc.Send "deldataxref:" & offset & ":" & tova
 End Sub
 
Function FuncVAByName(name)
    FuncVAByName = ipc.SendCmdRecvInt("name_va:" & name)
End Function

Function RenameFunc(oldname, newName) As Boolean
    RenameFunc = IIf(ipc.SendCmdRecvInt("rename:" & oldname & ":" & newName) = 1, True, False)
End Function



Function Jump(va As Long)
    'ipc.Send "jmp:" & va
    ipc.QuickCall qcmJmpAddr, va
End Function

Function JumpRVA(rva As Long)
    ipc.QuickCall qcmJmpRVA, rva
End Function

Function refresh()
    'ipc.Send "refresh:"
    ipc.QuickCall qcmRefresh
End Function

Function Undefine(offset)
    'ipc.Send "undefine:" & offset
    ipc.QuickCall qcmUndef, CLng(offset)
End Function

Function ShowEA(offset)
    'ipc.Send "show:" & offset
    ipc.QuickCall qcmShow, CLng(offset)
End Function

Function HideEA(offset)
    'ipc.Send "hide:" & offset
    ipc.QuickCall qcmHide, CLng(offset)
End Function

 Sub RemoveName(offset)
    'ipc.Send "remname:" & offset
    ipc.QuickCall qcmRemName, CLng(offset)
 End Sub

 
Sub MakeCode(offset)
    'ipc.Send "makecode:" & offset
    ipc.QuickCall qcmMakeCode, CLng(offset)
End Sub

Function FuncIndexFromVA(va)
    'FuncIndexFromVA = ipc.SendCmdRecvInt("funcindex:" & va)
    FuncIndexFromVA = ipc.QuickCall(qcmFuncIdx, CLng(va))
End Function

Function NextEA(va)
    'NextEA = ipc.SendCmdRecvInt("nextea:" & va)
    NextEA = ipc.QuickCall(qcmNextEa, CLng(va))
End Function

Function PrevEA(va)
    'PrevEA = ipc.SendCmdRecvInt("prevea:" & va)
    PrevEA = ipc.QuickCall(qcmPrevEa, CLng(va))
End Function

 Function NumFuncs() As Long
    'NumFuncs = ipc.SendCmdRecvInt("numfuncs")
    NumFuncs = ipc.QuickCall(qcmNumFuncs)
 End Function

 Function FunctionStart(functionIndex)
    'FunctionStart = ipc.SendCmdRecvInt("funcstart:" & functionIndex)
    FunctionStart = ipc.QuickCall(qcmFuncStart, CLng(functionIndex))
 End Function

 Function FunctionEnd(functionIndex)
'    FunctionEnd = ipc.SendCmdRecvInt("funcend:" & functionIndex)
    FunctionEnd = ipc.QuickCall(qcmFuncEnd, CLng(functionIndex))
 End Function
  
 Function ReadByte(va)
'    ReadByte = CByte(ipc.SendCmdRecvInt("readbyte:" & va))
    ReadByte = CByte(ipc.QuickCall(qcmReadByte, CLng(va)))
End Function

Function OriginalByte(va)
    'OriginalByte = CByte(ipc.SendCmdRecvInt("orgbyte:" & va))
    OriginalByte = CByte(ipc.QuickCall(qcmOrgByte, CLng(va)))
End Function

Function ImageBase() As Long
    'ImageBase = ipc.SendCmdRecvInt("imgbase")
    ImageBase = ipc.QuickCall(qcmImgBase)
End Function

Function ScreenEA() As Long
    ScreenEA = ipc.QuickCall(qcmScreenEA)
End Function

'old style calls...
    'FuncVAByName = ipc.SendCmdRecvInt("name_va:" & name & ":" & Form1.hwnd) 'old style with data callback..
    'FuncIndexFromVA = ipc.SendCmdRecvInt("funcindex:" & va & ":" & Form1.hwnd) 'old style with data callback..
    'NextEA = ipc.SendCmdRecvInt("nextea:" & va & ":" & Form1.hwnd) 'old style with data callback..
    'PrevEA = ipc.SendCmdRecvInt("prevea:" & va & ":" & Form1.hwnd) 'old style with data callback..
    'NumFuncs = ipc.SendCmdRecvInt("numfuncs:" & Form1.hwnd) 'old style with data callback..
    'FunctionStart = ipc.SendCmdRecvInt("funcstart:" & functionIndex & ":" & Form1.hwnd) 'old style with data callback..
    'FunctionEnd = ipc.SendCmdRecvInt("funcend:" & functionIndex & ":" & Form1.hwnd) 'old style with data callback..
    'x = ipc.SendCmdRecvInt("readbyte:" & va & ":" & Form1.hwnd) 'old style with data callback..
    'x = ipc.SendCmdRecvInt("orgbyte:" & va & ":" & Form1.hwnd) 'old style with data callback..
    'x = ipc.SendCmdRecvInt("imgbase:" & Form1.hwnd) 'old style with data callback..



' Sub DelFunc(addr)
'    Module1.DelFunc CLng(addr)
' End Sub
'
'Sub AddProgramComment(cmt)
'   Module1.AddProgramComment CStr(cmt)
'End Sub
'
' Function ScreenEA()
'    ScreenEA = Module1.ScreenEA
' End Function
'


'Function GetAsmBlock(start, leng)
'    GetAsmBlock = GetAsmRange(CLng(start), CLng(leng))
'End Function
'
'
'Function GetBytes(start, leng)
'    Dim b() As Byte, ret(), i As Integer
'
'    ReDim b(leng)
'    ReDim ret(leng)
'
'    Module1.GetBytes CLng(start), b(0), CLng(leng)
'
'    For i = 0 To leng
'        ret(i) = b(i)
'    Next
'
'    GetBytes = ret()
'
'End Function
'
'

'
'
'Sub AnalyzeArea(startat, endat)
'    Module1.AnalyzeArea CLng(startat), CLng(endat)
'End Sub
'

Function AskValue(Optional prompt, Optional defVal) As String
    AskValue = InputBox(prompt, , defVal)
End Function

Sub Exec(cmd)
   On Error Resume Next
   Shell cmd, vbNormalFocus
End Sub

Function ReadFile(filename) As Variant
  Dim f As Long
  Dim Temp As Variant
  f = FreeFile
  Temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     Temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = Temp
End Function

Sub WriteFile(path As String, it As Variant)
    Dim f As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

Sub AppendFile(path, it)
    Dim f As Long
    f = FreeFile
    Open path For Append As #f
    Print #f, it
    Close f
End Sub

Function FileExists(path As String) As Boolean
  On Error GoTo hell
  Dim tmp As String
  tmp = Replace(path, "'", Empty)
  tmp = Replace(tmp, """", Empty)
  If Len(tmp) = 0 Then Exit Function
  If Dir(tmp, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  Exit Function
hell: FileExists = False
End Function

Function DeleteFile(fpath As String) As Boolean
 On Error GoTo hadErr
    Kill fpath
    DeleteFile = True
 Exit Function
hadErr:
'MsgBox "DeleteFile Failed" & vbCrLf & vbCrLf & fpath
DeleteFile = False
End Function

Function getClipboard()
    getClipboard = Clipboard.GetText()
End Function

Function setClipboard(x)
    Clipboard.Clear
    Clipboard.SetText x
End Function

Function OpenFileDialog() As String
    OpenFileDialog = dlg.OpenDialog(AllFiles)
End Function

Function SaveFileDialog() As String
    SaveFileDialog = dlg.SaveDialog(AllFiles)
End Function

Function BenchMark() As Long
    Dim i As LARGE_INTEGER
    QueryPerformanceCounter i
    BenchMark = i.lowpart
End Function

'Function refListToArray(x) As Long() 'no way to use a vb6 array in js?
'    On Error Resume Next
'    Dim i As Long, y() As Long
'    x = Split(x, ",")
'    ReDim y(UBound(x) - 1)
'    For i = 0 To UBound(x) - 1
'        y(i) = CLng(x(i))
'    Next
'    refListToArray = y
'End Function
 
