VERSION 5.00
Begin VB.Form frmVBDeclares 
   Caption         =   "VB API Declare Name Extractor (Names thunks in IDA)"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   4215
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmVBDeclares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Text2 As String

Private Sub cmdSave_Click()
    Dim target As String
    Dim base As String
    
    target = Text2
    If Not FileExists(target) Then
        MsgBox "target not found!"
        Exit Sub
    End If
    
    If Len(Text1) = 0 Then
        MsgBox "Nothing to save"
        Exit Sub
    End If
    
    target = target & ".idc"
    WriteFile target, Text1
    
    MsgBox "Saved as: " & vbCrLf & vbCrLf & target, vbInformation

End Sub
    
    
Sub Initialize()
    
    
    Dim pe As New CPEOffsets
    Dim sig As String
    Dim offsets()
    Dim fva As Long
    Dim apiVa As Long
    Dim va As Long
    Dim X As String
    Dim p() As Byte
    Dim apiName As String
    
    Dim tmp As Long
    Dim target As String
    
    Dim curFile As String
    
    Text2 = loadedFile
    
    target = Text2
    If Not FileExists(target) Then
        MsgBox "target not found!"
        Unload Me
    End If
    
   
    
    sig = Chr(&HB) & Chr(&HC0) & Chr(&H74) & Chr(2) & Chr(&HFF) & Chr(&HE0) & Chr(&H68)
    
    push offsets, "#define UNLOADED_FILE   1"
    push offsets, "#include <idc.idc>"
    push offsets, ""
    push offsets, "static main(void) {"
    
    X = ReadFile(target)
    p = StrConv(X, vbFromUnicode)
    
    If Not pe.LoadFile(target) Then
        MsgBox "Could not load: " & target & vbCrLf & vbCrLf & "Error:" & pe.errMessage
        Unload Me
    End If

    i = InStr(X, sig)
    While i > 0
        va = pe.ImageBase + pe.OffsetToRVA(i)
        fva = va - 6
        apiVa = GetLongFromFileOffset(X, i + Len(sig))  'VA offset of libraryname struc
        apiVa = pe.RvaToOffset(apiVa - pe.ImageBase)
        apiVa = GetLongFromFileOffset(X, apiVa + 5) 'Extract VA offset of apiname
        apiVa = pe.RvaToOffset(apiVa - pe.ImageBase)
        apiName = ScanString(p, apiVa)
        push offsets, vbTab & "MakeName(0x" & Hex(fva) & ",""" & apiName & """);"
        i = InStr(i + 1, X, sig)
    Wend
    
    push offsets, "}"
    Text1 = Join(offsets, vbCrLf)
    
End Sub

 

Function ScanString(p() As Byte, ByVal fOffset As Long) As String
    Dim ret As String
    Dim b As Byte
    
    b = p(fOffset)
    While b <> 0
        ret = ret & Chr(b)
        fOffset = fOffset + 1
        b = p(fOffset)
    Wend
    
    ScanString = ret
    
    
    
End Function

Function GetLongFromFileOffset(X As String, fOffset As Long) As Long
   Dim tmp As String
   Dim ltmp() As Byte
   Dim retVal As Long
   
    tmp = Mid(X, fOffset, 4)
    ltmp = StrConv(tmp, vbFromUnicode)
    CopyMemory retVal, ltmp(0), 4
    GetLongFromFileOffset = retVal
        
End Function

Function ReadFile(filename)
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function



Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub



Private Sub Text2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text2 = Data.Files(1)
End Sub


Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function



Function GetParentFolder(path) As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
End Function


Function FileNameFromPath(fullpath) As String
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function



Sub WriteFile(path, it)
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

'.text:00405F64 LibraryName27   db 'kernel32.dll',0     ; DATA XREF: .text:aLibraryName27o
'.text:00405F64                                         ; .text:aLibraryName26o ...
'.text:00405F71                 db    0
'.text:00405F72                 db    0
'.text:00405F73                 db    0
'.text:00405F74                 db    6
'.text:00405F75                 db    0
'.text:00405F76                 db    0
'.text:00405F77                 db    0
'.text:00405F78 LibraryFunction27 db 'Sleep',0          ; DATA XREF: .text:00405F84o
'.text:00405F7E                 db    0
'.text:00405F7F                 db    0
'.text:00405F80 aLibraryName27  dd offset LibraryName27 ; DATA XREF: .text:00405750o
'.text:00405F80                                         ; sub_405F98:loc_405FA3o
'.text:00405F80                                         ; "kernel32.dll"
'.text:00405F84                 dd offset LibraryFunction27 ; "Sleep"
'.text:00405F88                 dd 40000h               ; word1x08
'.text:00405F88                 dd offset unk_42E558    ; offset1_0x0c
'.text:00405F88                 dd 0                    ; word2_0x10
'.text:00405F88                 dd 0                    ; word3_0x14
'.text:00405F98
'.text:00405F98 ; ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦ S U B R O U T I N E ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
'.text:00405F98
'.text:00405F98
'.text:00405F98 sub_405F98      proc near               ; CODE XREF: .text:0040E680p
'.text:00405F98                                         ; sub_40FB1C+5F9p ...
'.text:00405F98                 mov     eax, dword_42E560
'.text:00405F9D                 or      eax, eax
'.text:00405F9F                 jz      short loc_405FA3
'.text:00405FA1                 jmp     eax
'.text:00405FA3 ; ---------------------------------------------------------------------------
'.text:00405FA3
'.text:00405FA3 loc_405FA3:                             ; CODE XREF: sub_405F98+7j
'.text:00405FA3                 push    offset aLibraryName27
'.text:00405FA8                 mov     eax, offset DllFunctionCall
'.text:00405FAD                 call    eax ; DllFunctionCall
'.text:00405FAF                 jmp     eax
'.text:00405FAF sub_405F98      endp
'
'
'
'405F98 : a1 60 e5 42  0                mov     eax, dword_42E560
'405 F9D:  b c0 Or eax, eax
'405F9F : 74  2                         jz      short loc_405FA3
'405FA1 : ff e0                         jmp     eax
'405FA3 : 68 80 5f 40  0                push    offset aLibraryName27
'405FA8 : b8 70 4f 40  0                mov     eax, offset DllFunctionCall
'405FAD : ff d0                         call    eax ; DllFunctionCall
'
'
'

