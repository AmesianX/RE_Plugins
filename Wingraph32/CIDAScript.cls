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

Public WithEvents ipc As CIpc
Attribute ipc.VB_VarHelpID = -1

Private m_listen_hwnd As Long
Private m_ipcName As String
Private m_initilized As Boolean

Property Get Initilized() As Boolean
    Initilized = m_initilized
End Property

Function Initialize(hwnd As Long, ipcName As String)
    
    If m_initilized Then Exit Function
    
    Set ipc = New CIpc
    m_ipcName = ipcName
    m_listen_hwnd = hwnd
    ipc.FindClient "IDA_SERVER"
    ipc.Listen ipcName, hwnd
    
    m_initilized = True
    
End Function

Property Get isUp() As Boolean
    isUp = ipc.FindClient("IDA_SERVER")
End Property

Function alert(msg)
    If IsArray(msg) Then
        MsgBox "Array: " & vbCrLf & vbCrLf & Join(msg, vbCrLf)
    Else
        MsgBox msg
    End If
End Function

Function Jump(va As Long)
    ipc.Send "jmp:" & va
End Function

Function JumpName(name As String)
    ipc.Send "jmp_name:" & name
End Function

Function Message(msg As String)
    ipc.Send "msg:" & msg
End Function

Property Get LoadedFile() As String
    LoadedFile = ipc.SendAndRecv("loadedfile:" & m_listen_hwnd)
End Property

Property Get ImageBase() As Long
    On Error Resume Next
    Dim x
    x = ipc.SendAndRecv("imgbase:" & m_listen_hwnd)
    ImageBase = CLng(x)
End Property

Sub t(data As String)
    Form1.Text1.Text = Form1.Text1.Text & vbCrLf & data
    Form1.Text1.SelStart = Len(Form1.Text1.Text)
End Sub

Sub ClearLog()
    Form1.Text1.Text = Empty
End Sub

Function PatchByte(va, newVal)
    If CInt(newVal) > 255 Then Err.Raise 1, "PatchByte", "byte value must be 0-254"
    ipc.Send "patchbyte:" & va & ":" & newVal
    Refresh
End Function

Function ReadByte(va)
    Dim x
    x = ipc.SendAndRecv("readbyte:" & va & ":" & m_listen_hwnd)
    If Len(x) > 2 Then x = Right(x, 2)
    ReadByte = CByte(CInt("&h" & x))
End Function

Function OriginalByte(va)
    Dim x
    x = ipc.SendAndRecv("orgbyte:" & va & ":" & m_listen_hwnd)
    If Len(x) > 2 Then x = Right(x, 2)
    OriginalByte = CByte(CInt("&h" & x))
End Function

Function intToHex(x)
    On Error Resume Next
    intToHex = Hex(x)
    If Len(intToHex) = 1 Then intToHex = "0" & intToHex
    'intToHex = "\x" & intToHex
End Function

Function Refresh()
    ipc.Send "refresh:"
End Function

 Function NumFuncs() As Long
    NumFuncs = ipc.SendAndRecv("numfuncs:" & m_listen_hwnd)
 End Function

 Function FunctionStart(functionIndex)
    FunctionStart = ipc.SendAndRecv("funcstart:" & functionIndex & ":" & m_listen_hwnd)
 End Function

 Function FunctionEnd(functionIndex)
    FunctionEnd = ipc.SendAndRecv("funcend:" & functionIndex & ":" & m_listen_hwnd)
 End Function
 
 Function FunctionName(functionIndex)
    FunctionName = ipc.SendAndRecv("funcname:" & functionIndex & ":" & m_listen_hwnd)
 End Function

Function GetAsm(va)
    GetAsm = ipc.SendAndRecv("getasm:" & va & ":" & m_listen_hwnd)
End Function

Function InstSize(offset)
    'there has to be a real way to do this
    Dim x As String, tmp As String, i As Long, n As String
    Dim firstea As Long, secondea As Long
    Dim leng As Long
    
    leng = 40
    firstea = 0
    secondea = 0
    For i = 0 To leng - 1
        tmp = GetAsm(offset + i)
        If tmp = "Fail" Then tmp = Empty
        If Len(tmp) > 0 Then
            If firstea = 0 Then
                firstea = offset + i
            ElseIf secondea = 0 Then
                 secondea = offset + i
            End If
            If firstea > 0 And secondea > 0 Then Exit For
        End If
    Next
    
    InstSize = secondea - firstea
    
End Function

Function GetRefsTo(offset) 'returns CSV list
    Dim tmp
    tmp = ipc.SendAndRecvTillDone("refsto:" & offset & ":" & m_listen_hwnd)
    'tmp = Split(tmp, ",") javascript cant deal with string() must be variant()?
    GetRefsTo = tmp
End Function

Function GetRefsFrom(offset) 'returns CSV list
    Dim tmp
    tmp = ipc.SendAndRecvTillDone("refsfrom:" & offset & ":" & m_listen_hwnd)
    'tmp = Split(tmp, ",") javascript cant deal with string() must be variant()?
    GetRefsFrom = tmp
End Function

Function Undefine(offset)
    ipc.Send "undefine:" & offset
End Function

Function GetName(offset)
    GetName = ipc.SendAndRecv("getname:" & offset & ":" & m_listen_hwnd)
End Function

Function ShowEA(offset)
    ipc.Send "show:" & offset
End Function

Function HideEA(offset)
    ipc.Send "hide:" & offset
End Function
'
Function HideBlock(offset, leng)
    Dim x, i As Long
    x = offset + leng
    For i = offset To x
        HideEA i
    Next
    Refresh
End Function

Function ShowBlock(offset, leng)
    Dim x, i As Long
    x = offset + leng
    For i = offset To x
        ShowEA i
    Next
    Refresh
End Function


 Sub RemoveName(offset)
    ipc.Send "remname:" & offset
 End Sub

 Sub Setname(offset, name)
    ipc.Send "setname:" & offset & ":" & CStr(name)
 End Sub

Function Rename(curName, newName) As Boolean
    Dim ret As Long
    ret = ipc.SendAndRecv("rename:" & curName & ":" & newName & ":" & m_listen_hwnd)
    If ret = 1 Then Rename = True
End Function
 
Sub MakeCode(offset)
    ipc.Send "makecode:" & offset
End Sub

Sub AddComment(offset, comment)
    ipc.Send "addcomment:" & offset & ":" & comment
End Sub

Function GetComment(offset)
    GetComment = ipc.SendAndRecv("getcomment:" & offset & ":" & m_listen_hwnd)
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

'add these to IDA_Jscript
'-----------------------------------
Function FuncIndexFromVA(va)
    FuncIndexFromVA = ipc.SendAndRecv("funcindex:" & va & ":" & m_listen_hwnd)
End Function

Function FuncVAByName(name)
    FuncVAByName = ipc.SendAndRecv("name_va:" & name & ":" & m_listen_hwnd)
End Function

Function NextEA(va)
    NextEA = ipc.SendAndRecv("nextea:" & va & ":" & m_listen_hwnd)
End Function
'-----------------------------------



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


Sub Exec(cmd)
   On Error Resume Next
   Shell cmd, vbNormalFocus
End Sub

Function ReadFile(filename) As Variant
  Dim f As Long
  Dim temp As Variant
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
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

