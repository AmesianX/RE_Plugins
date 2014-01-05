Attribute VB_Name = "Module1"
Global fso As New CFileSystem2
Global dlg As New clsCmnDlg

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
 
Option Compare Binary

Sub setpb(cur, max)
    On Error Resume Next
    Form1.pb.Value = (cur / max) * 100
    Form1.Refresh
    DoEvents
 End Sub
 
Function FindMZOffsets(fpath As String) As Collection
    
    Dim f As Long, pointer As Long
    Dim buf()  As Byte
    Dim x As Long
    Dim fs As Long
    Dim ret As New Collection
     
    f = FreeFile
    
    If Not fso.FileExists(fpath) Then GoTo done
       
    ReDim buf(9000)
    Open fpath For Binary Access Read As f
    
    Form1.pb.Value = 0
    Do While pointer < LOF(f)
        pointer = Seek(f)
        x = LOF(f) - pointer
        If x < 1 Then Exit Do
        If x < 9000 Then ReDim buf(x)
        Get f, , buf()
        search buf, pointer, ret
        setpb pointer, LOF(f)
    Loop
    Form1.pb.Value = 0
    
    Close f
     
done:
    Set FindMZOffsets = ret
      
End Function

Private Sub search(buf() As Byte, offset As Long, c As Collection)
    
    Dim b As String
    Dim a As Long
    a = 1
    b = StrConv(buf, vbUnicode)
    
    Do
        a = InStr(a, b, "MZ")
        If a > 0 Then
            c.Add offset + a - 2
            a = a + 1
        End If
        
        If a >= Len(b) Then a = 0
        
    Loop While a > 0
    
    
End Sub

