Attribute VB_Name = "Module1"
 Public Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Global dlg As New clsCmnDlg
Global fso As New CFileSystem2

Sub FormPos(fform As Form, Optional andSize As Boolean = False, Optional save_mode As Boolean = False)
    
    On Error Resume Next
    
    Dim f, sz
    f = Split(",Left,Top,Height,Width", ",")
    
    If fform.WindowState = vbMinimized Then Exit Sub
    If andSize = False Then sz = 2 Else sz = 4
    
    For i = 1 To sz
        If save_mode Then
            ff = CallByName(fform, f(i), VbGet)
            SaveSetting App.EXEName, fform.name & ".FormPos", f(i), ff
        Else
            def = CallByName(fform, f(i), VbGet)
            ff = GetSetting(App.EXEName, fform.name & ".FormPos", f(i), def)
            CallByName fform, f(i), VbLet, ff
        End If
    Next
    
End Sub

Sub SaveMySetting(key, Value)
    SaveSetting App.EXEName, "Settings", key, Value
End Sub

Function GetMySetting(key, Optional defaultval = "")
    GetMySetting = GetSetting(App.EXEName, "Settings", key, defaultval)
End Function

Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = c(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function
