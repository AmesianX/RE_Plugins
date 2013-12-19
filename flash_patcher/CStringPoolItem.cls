VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStringPoolItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Index As Long
Public Offset As Long
Public RawOffset As Long
Public Data As String
Public DataLength As Long

Function LoadSelf(proto, idx, block As CAbcBlock)
    
    Index = idx
    proto = Trim(proto)
    proto = Replace(Replace(proto, vbCr, ""), vbLf, "")
    
    a = InStr(proto, ":")
    If a > 0 Then
        b = Mid(proto, 1, a - 1)
        b = Trim(Replace(b, "offset", ""))
        RawOffset = CLng(b)
        Offset = block.CodeOffset + RawOffset
    End If
    
    On Error Resume Next
    Data = Trim(Mid(proto, a + 1))
    DataLength = Len(Data)
    
End Function
