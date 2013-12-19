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
Attribute VB_Exposed = True
Option Explicit
    
Function NextAddr(offset)
     NextAddr = Module1.NextAddr(offset)
End Function

Function PrevAddr(offset)
    PrevAddr = Module1.PrevAddr(offset)
End Function

    
Function ShowEA(offset)
    Module1.ShowEA offset
End Function

Function HideEA(offset)
    Module1.HideEA offset
End Function

Function HideX(offset, leng)
    Dim x, i As Long
    x = offset + leng
    For i = offset To x
        Module1.HideEA i
    Next
    Refresh
End Function

Function ShowX(offset, leng)
    Dim x, i As Long
    x = offset + leng
    For i = offset To x
        Module1.ShowEA i
    Next
    Refresh
End Function
 
Function HideOffsets(startea, endea)
    Dim i As Long
    For i = startea To endea
        Module1.HideEA i
    Next
    Refresh
End Function
 
Function ShowOffsets(startea, endea)
    Dim i As Long
    For i = startea To endea
        Module1.ShowEA i
    Next
    Refresh
End Function

 
 
 Function FuncIndex(addr)
    FuncIndex = Module1.FuncIndex(CLng(addr))
 End Function
 
 Function FuncArgSize(index)
    FuncArgSize = Module1.FuncArgSize(CLng(index))
 End Function
 
 Function FuncColor(index) As Byte
    FuncColor = Module1.FuncColor(CLng(index))
 End Function
 
 Sub PatchByte(addr, valu)
    Module1.PatchByte CLng(addr), CByte(valu)
 End Sub
 
 Sub PatchWord(addr, valu)
    Module1.PatchWord CLng(addr), CLng(valu)
 End Sub
 
 Sub DelFunc(addr)
    Module1.DelFunc CLng(addr)
 End Sub
 
Sub AddComment(cmt, clr)
    Module1.AddComment CStr(cmt)
 End Sub
 
Sub AddProgramComment(cmt)
   Module1.AddProgramComment CStr(cmt)
End Sub
 
 Sub AddCodeXRef(start, endd)
    Module1.AddCodeXRef CLng(start), CLng(endd)
 End Sub
 
 Sub DelCodeXRef(start, endd)
    Module1.DelCodeXRef CLng(start), CLng(endd)
 End Sub
 
 Sub AddDataXRef(start, endd)
    Module1.AddDataXRef CLng(start), CLng(endd)
 End Sub
 
 Sub DelDataXRef(start, endd)
    Module1.DelDataXRef CLng(start), CLng(endd)
 End Sub
 
 Sub MessageUI(Msg)
    Module1.MessageUI CStr(Msg)
 End Sub
 
 Function NumFuncs()
    NumFuncs = Module1.NumFuncs()
 End Function
 
 Function FunctionStart(functionIndex)
    FunctionStart = Module1.FunctionStart(CLng(functionIndex))
 End Function
 
 Function FunctionEnd(functionIndex)
    FunctionEnd = Module1.FunctionEnd(CLng(functionIndex))
 End Function
  
 Sub Jump(offset)
    Module1.Jump CLng(offset)
 End Sub
 
 Sub RemvName(offset)
    Module1.RemvName CLng(offset)
 End Sub
 
 Sub Setname(offset, name)
    Module1.Setname CLng(offset), CStr(name)
 End Sub
 
 Sub Refresh()
    Module1.aRefresh
 End Sub
 
 Function ScreenEA()
    ScreenEA = Module1.ScreenEA
 End Function
 
 Sub SelBounds(selstart, selend)
    Module1.SelBounds CLng(selstart), CLng(selend)
 End Sub
 
 Function selstart()
    Dim a, b
    SelBounds a, b
    selstart = a
 End Function
 
 Function selend()
    Dim a, b
    SelBounds a, b
    selend = b
 End Function
 
 Function GetFuncName(offset)
    GetFuncName = GetFName(CLng(offset))
 End Function
 
 Function GetAsm(offset)
    GetAsm = GetAsmCode(offset)
 End Function

Function GetAsmBlock(start, leng)
    GetAsmBlock = GetAsmRange(CLng(start), CLng(leng))
End Function


Function GetBytes(start, leng)
    Dim b() As Byte, ret(), i As Integer
    
    ReDim b(leng)
    ReDim ret(leng)
    
    Module1.GetBytes CLng(start), b(0), CLng(leng)
    
    For i = 0 To leng
        ret(i) = b(i)
    Next
    
    GetBytes = ret()
    
End Function

Function LogIt(Msg)
    Dim t As TextBox
    Set t = frmIDAScript.txtLog
    
    t.selstart = t.SelLength
    t.SelText = Msg & vbCrLf
    
End Function

Sub ClearLog()
    frmIDAScript.txtLog.Text = Empty
End Sub

Sub Done()
    frmPluginSample.Done
End Sub

Sub MakeCode(offset)
    Module1.MakeCode CLng(offset)
End Sub

Sub Undefine(offset)
    Module1.Undefine CLng(offset)
End Sub

Sub AnalyzeArea(startat, endat)
    Module1.AnalyzeArea CLng(startat), CLng(endat)
End Sub

Function GetName(offset)
    GetName = Module1.GetName(CLng(offset))
End Function

Function OrgByte(offset)
    OrgByte = Module1.OriginalByte(CLng(offset))
End Function

Sub Exec(cmd)
    On Error Resume Next
    Shell cmd, vbNormalFocus
End Sub


Sub SetComment(offset, comment)
    On Error Resume Next
    Module1.SetComment CLng(offset), CStr(comment)
End Sub

Function GetComment(offset)
    GetComment = Module1.Get_Comment(offset)
End Function


