Attribute VB_Name = "Module1"

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
 
Global p As New CParser
Global dlg As New clsCmnDlg2


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

Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
init: ReDim ary(0): ary(0) = Value
End Sub

Function FolderExists(path As String) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbDirectory) <> "" Then FolderExists = True
End Function

Function GetParentFolder(path) As String
    Dim tmp() As String
    Dim ub As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
End Function

Function GetBaseName(path As String) As String
    Dim tmp() As String
    Dim ub As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetBaseName = Mid(ub, 1, InStrRev(ub, ".") - 1)
    Else
       GetBaseName = ub
    End If
End Function

Sub WriteFile(path As String, it As Variant)
    Dim f As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
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

Function RandomNum() As Long
    Dim tmp As Long
    Dim tries As Long
    
    On Error GoTo again
tryit:

    Randomize
    tmp = Round(Timer * Now * Rnd(), 0)
    RandomNum = tmp
    
    Exit Function
again:
    
    If tries < 10 Then
        tries = tries + 1
        GoTo tryit
    End If
    
End Function

Function GetFreeFileName(ByVal folder As String, Optional extension = ".txt") As String
    
    On Error GoTo handler 'can have overflow err once in awhile :(
    Dim i As Integer
    Dim tmp As String

    If Not FolderExists(folder) Then Exit Function
    If Right(folder, 1) <> "\" Then folder = folder & "\"
    If Left(extension, 1) <> "." Then extension = "." & extension
    
again:
    Do
      tmp = folder & RandomNum() & extension
    Loop Until Not FileExists(tmp)
    
    GetFreeFileName = tmp
    
Exit Function
handler:

    If i < 10 Then
        i = i + 1
        GoTo again
    End If
    
End Function

Public Function GetShortName(sFile As String) As String
    Dim sShortFile As String * 300
    Dim lResult As Long

    If Not FileExists(sFile) Then
        MsgBox "GetshortName file must exist to work..: " & sFile
        GetShortName = sFile
        Exit Function
    End If
    
    'file must exist or this will fail...
    lResult = GetShortPathName(sFile, sShortFile, Len(sShortFile))

    'Trim out unused characters from the string.
    GetShortName = Left$(sShortFile, lResult)

    If Not FileExists(GetShortName) Then GetShortName = sFile
End Function

Sub SetLiColor(li As ListItem, newcolor As Long)
    Dim f As ListSubItem
'    On Error Resume Next
    li.ForeColor = newcolor
    For Each f In li.ListSubItems
        f.ForeColor = newcolor
    Next
End Sub

Sub LV_LastColumnResize(lv As ListView)
    On Error Resume Next
    lv.ColumnHeaders(lv.ColumnHeaders.Count).Width = lv.Width - lv.ColumnHeaders(lv.ColumnHeaders.Count).Left - 100
End Sub

Function isDecimalNumber(x) As Boolean
    
    'Debug.Print isDecimalNumber("32")    'true
    'Debug.Print isDecimalNumber("32 ")   'true
    'Debug.Print isDecimalNumber("232a ") 'false
    ' Stop
     
    On Error GoTo hell
    Dim l As Long
    
    For i = 1 To Len(x) - 1
        c = Mid(x, i, 1)
        If Not IsNumeric(c) Then Exit Function
    Next
    
    l = CLng(x)
    isDecimalNumber = True
    
hell:
    Exit Function
    
End Function

Function StringOpcodesToBytes(OpCodes) As Byte()
    
    'Debug.Print StrConv(StringOpcodesToBytes("41 42 43 44"), vbUnicode)
    'Stop
    
    On Error Resume Next
    Dim b() As Byte
    
    tmp = Split(Trim(OpCodes), " ")
    ReDim b(UBound(tmp))
    
    For i = 0 To UBound(tmp)
        b(i) = CByte(CInt("&h" & tmp(i)))
    Next
    
    StringOpcodesToBytes = b()
    
End Function

Function pad(x, Optional sz = 8)
    a = Len(x) - sz
    If a < 0 Then
        pad = x & Space(Abs(a))
    Else
        pad = x
    End If
End Function

Function objKeyExistsInCollection(c As Collection, k As String) As Boolean
    On Error GoTo hell
    Set x = c(k)
    objKeyExistsInCollection = True
hell:
End Function
