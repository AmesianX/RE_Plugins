VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBinWrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public isOpen As Boolean
Public fhand As Long

Sub fOpen(pth)
    fhand = FreeFile
    Open pth For Binary As fhand
End Sub

Sub fClose()
    Close fhand
    fhand = 0
End Sub

Sub PutC(b)
    Put fhand, , CByte(b)
End Sub

Function GetC()
    Dim b As Byte
    Get fhand, , b
    GetC = b
End Function

Function FLen()
    FLen = LOF(fhand)
End Function

Sub PutCO(offset, b)
    Put fhand, offset, CByte(b)
End Sub

Function GetCO(offset)
    Dim b As Byte
    Get fhand, offset, b
    GetCO = b
End Function

Sub PutStr(offset, str)
    Put fhand, offset, str
End Sub


Function GetAry(offset, size)
    Dim b() As Byte
    ReDim b(size - 1)
    
    Get fhand, offset, b()
    
    Dim v(), i
    ReDim v(size - 1)
    
    For i = 0 To size - 1
        v(i) = b(i)
    Next
    
    GetAry = v()
    
End Function

Function PutAry(offset, ary)
    
    Dim i As Integer
    
    Put fhand, offset, CByte(ary(0))
    
    For i = 1 To UBound(ary)
         Put fhand, , CByte(ary(i))
    Next
    
End Function



