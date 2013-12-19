VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public curFile As String
Public Sections As New Collection
Public ImageBase As Long
Public isLoaded As Boolean

Sub FilloutListView(lv As Object)
        
    If Not isLoaded Then
        MsgBox "Sections not loaded yet"
        Exit Sub
    End If
    
    Dim cs As CSection, li As ListItem
    
    lv.ListItems.Clear
    
    For Each cs In Sections
        Set li = lv.ListItems.Add(, , cs.nameSec)
        li.SubItems(1) = Hex(cs.VirtualAddress)
        li.SubItems(2) = Hex(cs.VirtualSize)
        li.SubItems(3) = Hex(cs.PointerToRawData)
        li.SubItems(4) = Hex(cs.SizeOfRawData)
        li.SubItems(5) = Hex(cs.Characteristics)
    Next
    
    Dim i As Integer
    For i = 1 To lv.ColumnHeaders.Count
        lv.ColumnHeaders(i).Width = 1000
    Next
    With lv.ColumnHeaders(i - 1)
        .Width = lv.Width - .Left - 100
    End With
    
    
End Sub


Sub LoadSections(fpath As String)
    
    If Not FileExists(fpath) Then
        MsgBox "File not found: " & fpath
        Exit Sub
    End If
    
    curFile = fpath
    
    Dim f As Long, i As Long
    Dim dos As ImgDosHeader
    Dim opt As ImgOptHeader
    Dim udtSH As SECTION_HEADER
    Dim csect As CSection
    
    f = FreeFile
    
    Open fpath For Binary As f
    Get f, , dos
    Get f, (dos.pOptHeader + 1 + 52), ImageBase
    Get f, dos.pOptHeader + 1, opt
     
    Set Sections = New Collection
    
    For i = 1 To opt.NumberofSections
        Get f, , udtSH
        Set csect = New CSection
        csect.LoadStruc udtSH
        Sections.Add csect
    Next
           
    Close f
    
    isLoaded = True
    
End Sub

 
 
Function RvaToOffset(ByVal rva As Long, Optional ByRef VirtualSectionBase As Long, Optional sectName As String)
        
    If Sections.Count = 0 Then
        MsgBox "No Sections Loaded", vbInformation
        Exit Function
    End If
    
    Dim c As CSection
    Dim vBase As Long, vSize As Long, vMax As Long
    
    For Each c In Sections
       vBase = c.VirtualAddress
       vSize = c.VirtualSize
       vMax = vBase + vSize
       
       If rva >= vBase And rva < vMax Then 'right section
            rva = rva - vBase
            rva = rva + c.PointerToRawData
            RvaToOffset = rva
            VirtualSectionBase = vBase
            sectName = c.nameSec
            Exit Function
       End If
    Next
       
    
End Function

Function OffsetToRVA(ByVal fOffset As Long, Optional sectName As String)
    
    If Sections.Count = 0 Then
        MsgBox "No Sections Loaded", vbInformation
        Exit Function
    End If
    
    Dim c As CSection
    Dim rBase As Long, rSize As Long, rMax As Long
    
    For Each c In Sections
       rBase = c.PointerToRawData
       rSize = c.SizeOfRawData
       rMax = rBase + rSize
       
       If fOffset >= rBase And fOffset < rMax Then 'right section
            OffsetToRVA = c.VirtualAddress + fOffset - rBase
            sectName = c.nameSec
            Exit Function
       End If
    Next

End Function


Function GetSectNameFromBase(virtBase As Long) As String
    Dim c As CSection
    For Each c In Sections
        If c.VirtualAddress = virtBase Then
            GetSectNameFromBase = c.nameSec
            Exit Function
        End If
    Next
End Function
