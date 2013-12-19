VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIntelliLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type LongWords
    LoWord As String
    HiWord As String
End Type

Public WithEvents txt As TextBox
Attribute txt.VB_VarHelpID = -1
Public WithEvents lst As ListBox
Attribute lst.VB_VarHelpID = -1
Public tok As New Collection
Public proto As New Collection

Sub Load(Optional fpath As String)
    Dim ary, i, y, x
    On Error Resume Next
    
    Dim protof As String
    
    If Len(fpath) > 0 Then
        protof = fpath
    Else
        protof = App.path & "\protos.txt"
    End If
    
    If Not FileExists(protof) Then
        MsgBox "Prototype file not found", vbInformation
        Exit Sub
    End If
    
    ary = ReadFile(protof)
    ary = Split(ary, vbCrLf)
    
    For i = 0 To UBound(ary)
        x = Trim(ary(i))
        If Len(x) = 0 Then GoTo nextone
            
        If left(x, 1) = ";" Then GoTo nextone
        If left(x, 1) = "#" Then GoTo nextone
        If left(x, 1) = "'" Then GoTo nextone
        If left(x, 1) = "/" Then GoTo nextone
            
        x = Replace(x, "sub", Empty, , 1, vbTextCompare)
        x = Replace(x, "function", Empty, , 1, vbTextCompare)
        x = Replace(x, "property get", Empty, , 1, vbTextCompare)
        
        y = InStr(x, "(")
        If y > 1 Then x = Mid(x, 1, y - 1)
        
        x = Trim(x)
        
        tok.Add x, LCase(CStr(x))
        proto.Add ary(i), LCase(CStr(x))
        
nextone:
    Next
    
End Sub

Sub Search()
    Dim s, lw As String, t As String
    Dim ret As New Collection
    
    On Error Resume Next
    
    lw = Trim(LCase(lastWord))
    'frmIDAScript.Caption = lw
    
    For Each s In tok
        t = LCase(left(s, Len(lw)))
        If t = lw Then
            ret.Add s
        End If
    Next
    
    If ret.Count = 1 Then
        InsertToken ret(1)
    ElseIf ret.Count > 1 Then
        DisplayList ret
    Else
        DisplayList tok
    End If
    
    
End Sub

Private Sub DisplayList(col As Collection)
    On Error Resume Next
    Dim s
    
    lst.Clear
    PositionList

    For Each s In col
        lst.AddItem s
    Next
    
    ShowWindow lst.hwnd, 1
    lst.SetFocus
    
End Sub

Sub InsertToken(token As String)
    Dim lw
    On Error Resume Next
    lw = lastWord
    'MsgBox Len(lw) & " " & txt.selstart
    txt.selstart = txt.selstart - Len(lw)
    txt.SelLength = Len(lw)
    txt.SelText = Trim(token)
    txt.Parent.Caption = GetProto(token)
End Sub

Function isToken(token As String) As Boolean
    On Error GoTo hell
    Dim x
    x = tok(token)
    isToken = True
    Exit Function
hell:
    isToken = False
End Function

Function GetProto(token As String) As String
    On Error Resume Next
    GetProto = proto(LCase(token))
End Function

Function lastWord() As String
    On Error Resume Next
    Dim p
    
    While lastChar = " "
        txt.selstart = txt.selstart - 1
    Wend
    
    Dim lastLine
    
    p = InStrRev(txt.Text, vbCrLf, txt.selstart)
    If p < 1 Then p = 1
    lastLine = Mid(txt.Text, p + 2, txt.selstart - p + 2)
    
    p = InStr(lastLine, vbCr)
    If p > 0 Then lastLine = Mid(lastLine, 1, p - 1)
    lastLine = Trim(lastLine)
    
    p = InStrRev(lastLine, "(") + 1
    If p < 2 Then p = InStrRev(lastLine, " ") + 1
    If p < 2 Then p = 1
        
    If p > 0 Then
        lastWord = Mid(lastLine, p)
    End If

    lastWord = Trim(lastWord)
    lastWord = Replace(lastWord, vbTab, Empty)
     
    
End Function

Function lastChar() As String
    On Error Resume Next
    lastChar = Mid(txt.Text, txt.selstart - 1, 1)
End Function


Private Sub Lst_KeyPress(KeyAscii As Integer)
    
    On Error Resume Next
    
    Select Case KeyAscii
        Case 27 'escape
            ShowWindow lst.hwnd, 0
            txt.SetFocus
        Case 13, 20 'return
            InsertToken lst.List(lst.ListIndex)
            ShowWindow lst.hwnd, 0
            txt.SetFocus
    End Select
     
End Sub

Private Sub txt_Click()
    ShowWindow lst.hwnd, 0
End Sub

Private Sub txt_KeyPress(KeyCode As Integer)
    If KeyCode = 96 Then
        'MsgBox ct.lastWord
        Search
        KeyCode = 0
    End If
End Sub
















Private Function ReadFile(filename)
  Dim f, temp
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Private Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

Private Sub PositionList()
    Dim pos As Long, caret As Long
    Dim p As POINTAPI
    Const EM_POSFROMCHAR = &HD6
    
    Const leftOffset = 200
    Const topOffset = 75
    
    caret = txt.selstart - 1
    pos = SendMessage(txt.hwnd, EM_POSFROMCHAR, caret, 0)
    
    'pos is really a short of the x and a short of the y, packed into the
    'x is the lower 2 bytes, y is the upper 2 bytes.
    
    lst.left = txt.left + (LoWord(pos) * 15) + leftOffset
    lst.Top = txt.Top + (HiWord(pos) * 15) + topOffset
    
End Sub

Private Function HiWord(l As Long) As Integer
    HiWord = CInt("&h" & LongToWords(l).HiWord)
End Function

Private Function LoWord(l As Long) As Integer
    LoWord = CInt("&h" & LongToWords(l).LoWord)
End Function

Private Function LongToWords(l As Long) As LongWords
    Dim w(3) As Byte
    CopyMemory w(0), l, 4
    LongToWords.HiWord = Hexit(w(3)) & Hexit(w(2))
    LongToWords.LoWord = Hexit(w(1)) & Hexit(w(0))
End Function

Private Function Hexit(x) As String
    Hexit = Hex(x)
    If Len(Hexit) < 2 Then Hexit = "0" & Hexit
End Function
