VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAbcBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public StartAt As Long
Public EndAt As Long
Public CodeOffset As Long
Public Data As String
Public Index As Long

Private kWarned As Boolean
Public SHIFT_CONST_K As Long

Function LoadSelf()

    Dim f As CFunction
    
    tmp = Split(Data, vbCrLf & vbCrLf)
    
    p.dbg.AddItem UBound(tmp) & " functions found"
    
    'extract function blocks...
    For i = 1 To UBound(tmp) - 1
        Set f = New CFunction
        x = tmp(i)
        a = InStr(x, vbCrLf)
        b = Mid(x, 1, a)
        r = Mid(x, a)
        f.LoadSelf b, r, Me
        p.Functions.Add f
    Next
    
    'strign poolstarts with String Constant Pool Entries
    'string pool ends with Namespace Constant Pool Entries
    
    a = InStr(Data, "String Constant Pool Entries")
    If a > 0 Then
        b = InStrRev(Data, ":", a) + 1
        Count = CLng(Trim(Mid(Data, b, a - b)))
        p.dbg.AddItem Count & " string pool entries found"
        If Count > 0 Then
            b = InStr(a, Data, "Namespace Constant Pool Entries")
            If b > 0 Then
                tmp = Mid(Data, a, b - a)
                tmp = Split(tmp, vbCrLf)
                For i = 1 To UBound(tmp) - 1
                    Set spi = New CStringPoolItem
                    spi.LoadSelf tmp(i), i, Me
                    p.StringPool.Add spi
                Next
            End If
        End If
    End If
                    
    tmp = Split(Data, "Pool Entries")
    For i = 0 To UBound(tmp) - 1
        x = tmp(i)
        a = InStrRev(x, vbLf)
        If a > 0 Then
            p.dbg.AddItem Mid(x, a + 1) & " Pool Entries"
        End If
    Next
    
    
End Function




Function SetKConstant(relativeOffset As Long, startBytes) As Boolean
    
    Dim ff As Long, b() As Byte, b2() As Byte
    Dim StartAt As Long
    
    b2() = StringOpcodesToBytes(startBytes)
    ReDim b(200)
    
    ff = FreeFile
    StartAt = relativeOffset + Me.CodeOffset
    
    Open p.DecompressedSWF For Binary As ff
    Get ff, StartAt, b()
    Close ff
    
    For i = 0 To UBound(b) - UBound(b2)
        If b(i) = b2(0) Then
            For j = 0 To UBound(b2)
                If b(i + j) <> b2(j) Then Exit For
            Next
            If j = UBound(b2) + 1 Then
                SHIFT_CONST_K = i - 1
                CodeOffset = CodeOffset + SHIFT_CONST_K
                SetKConstant = True
                p.dbg.AddItem "SetKConstant = " & SHIFT_CONST_K
                Exit Function
            End If
        End If
    Next
        
    If Not kWarned Then
         p.dbg.AddItem "SetKConstant was not able to be set! could not locate " & startBytes & " near " & relativeOffset
         kWarned = True
    End If
    
End Function

