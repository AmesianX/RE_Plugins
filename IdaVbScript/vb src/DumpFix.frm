VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDumpFix 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IAT DumpFix Imports APIs from olly dump for CALL PTR and JMP IATs"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   1125
      TabIndex        =   4
      Top             =   3465
      Width           =   975
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   3480
      Width           =   975
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   3375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5953
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"DumpFix.frx":0000
   End
   Begin VB.CommandButton cmdParse 
      Caption         =   "Parse"
      Height          =   375
      Left            =   3660
      TabIndex        =   1
      Top             =   3420
      Width           =   1095
   End
   Begin VB.CommandButton cmdIntegrate 
      Caption         =   "Integrate"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   3420
      Width           =   1215
   End
End
Attribute VB_Name = "frmDumpFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this was made about as quick as they come
'does what it needs to, no frills

Dim unique As Collection
Dim loadedFile As String

Private Sub cmdSave_Click()
    On Error Resume Next
    writeFile loadedFile & ".idc", Text2.Text
End Sub

Private Sub cmdhelp_Click()
    Dim h As String
    
    h = "What does this do? " & vbCrLf & _
            "" & vbCrLf & _
            "This lets you rename IAT function pointers in IDA disasm easily " & vbCrLf & _
            "from data easily obtained from Olly. " & vbCrLf & _
            "" & vbCrLf & _
            "In Olly with target running and all modules loaded, right click on main" & vbCrLf & _
            "disasm window and search for all intermodular calls. Right click on this " & vbCrLf & _
            "table and choose to copy entire table to clipboard. (Make sure you are in" & vbCrLf & _
            "disasm of code segment!)" & vbCrLf & _
            "" & vbCrLf & _
            "Now paste this data to text file and save. Do memory dump, fix entry point," & vbCrLf & _
            "and disassemble with IDA. Now fir up this plugin and drag and drop imports" & vbCrLf & _
            "text file into this textbox. Hit parse and API names ansd addresses will be " & vbCrLf & _
            "extracted from data. Hit integrate, and it will add all of those names to those" & vbCrLf & _
            "addresses. " & vbCrLf & _
            "" & vbCrLf & _
            "Voila you have a readable disassembly again withouth the whoas of tryign to find" & vbCrLf & _
            "the right unpacker version or of manually rebuilding the import table. " & vbCrLf

    MsgBox h
    
End Sub

Private Sub cmdIntegrate_Click()
    On Error Resume Next
    
    Dim tmp() As String, x() As String
    Dim addr As Long, Name As String
    Dim i As Long
    
    tmp = Split(Text2.Text, vbCrLf)
    
    For i = 0 To UBound(tmp)
        tmp(i) = Trim(tmp(i))
        If Len(tmp(i)) = 0 Then GoTo nextone
        
        x = Split(tmp(i), ",")
        addr = CLng("&h" & x(0))
        Name = x(1)
        
        If addr > 0 Then Setname addr, Name
        
        'MsgBox i & ": " & addr & " : " & name & " : " & Err.Description
        
nextone:
    Next
    
    Call aRefresh
    Unload Me
    
End Sub

Private Sub cmdLoad_Click()
    
    On Error Resume Next
    
    Dim d As New clsCmnDlg
    Dim fso As New clsFileSystem
    
    Dim p As String
    p = d.OpenDialog(textFiles)
    If Len(p) = 0 Then Exit Sub
    Text2 = fso.ReadFile(p)
    
End Sub

Private Sub cmdParse_Click()
    On Error Resume Next
    
    Dim ret() As String
    
    header = Empty '"#define UNLOADED_FILE   1" & vbCrLf & _
             '"#include <idc.idc>" & vbCrLf & vbCrLf & _
             '"static main(void) {" & vbCrLf

    Set unique = New Collection
    
    f = Text2
    f = Split(f, vbCrLf) 'lines
    
    'MsgBox UBound(f)
    
    Dim addr As String
    Dim import As String
    
    For i = 0 To UBound(f)
              
        l = f(i)
        addr = Empty: import = Empty
        
        If InStr(1, l, "mutex", vbTextCompare) > 0 Then
            DoEvents 'to trap for break
        End If
        
        If InStr(l, "CALL") > 0 And InStr(l, "PTR") > 0 Then 'style 2
            ImportStyleCallPtr l, addr, import
        ElseIf InStr(l, "CALL") > 0 Then
            ImportStyleCall l, addr, import
        ElseIf InStr(l, "JMP") > 0 Then 'style 1
            ImportStyleJmp l, addr, import
        ElseIf InStr(l, ".") > 0 Then
            PointerTable l, addr, import
        End If
 
        
        If Len(import) = 0 Then GoTo nextone
        
        unique.Add CStr(import), CStr(import)
        
        'import = Replace(import, "-", "_") 'some chars are reserved for IDA names

        'MakeName(0X4010E8,  "THISISMYSUB_2");
        'tmp = tmp & vbTab & "MakeName(0X" & addr & ",""" & import & """);" & vbCrLf
        
        push ret, addr & "," & import
        
        
nextone:
    Next
    
    'Text2 = header & tmp & "}"
    
    Text2 = Join(ret, vbCrLf)
    
    
End Sub

Sub PointerTable(fileLine, addrVar, importNameVar)
'    43434394 7C91137A  ntdll.RtlDeleteCriticalSection
    l = Split(Trim(fileLine), " ")
    addrVar = l(0)
    importNameVar = l(UBound(l))
    
    a = InStr(importNameVar, ".")
    If a > 0 Then
        importNameVar = Mid(importNameVar, a + 1)
    End If
    
    If KeyExistsInCollection(unique, CStr(importNameVar)) Then
        importNameVar = Empty
        addrVar = Empty
    End If
    
End Sub


'all variables byref modificed here
Sub ImportStyleJmp(fileLine, addrVar, importNameVar)
    '00402A98  FF25 7CF14100  JMP DWORD PTR DS:[41F17C] ; ADVAPI32.AdjustTokenPrivileges
    '--------                                             ------------------------------
    l = Split(fileLine, " ") 'words (we want first(address) and last (api name)
    addrVar = l(0)
    importNameVar = l(UBound(l))
    
    a = InStr(importNameVar, ".")
    If a > 0 Then
        importNameVar = Mid(importNameVar, a + 1)
    End If
    
    If KeyExistsInCollection(unique, CStr(importNameVar)) Then
        importNameVar = Empty
        addrVar = Empty
    End If
    
End Sub




'all variables byref modificed here
Sub ImportStyleCallPtr(fileLine, addrVar, importNameVar)
    '00401000   CALL DWORD PTR DS:[405100]                KERNEL32.FreeConsole
    '                              ------                 --------------------
    
    l = Split(Trim(fileLine), " ") 'words (we want first(address) and last (api name)
    importNameVar = l(UBound(l))
    
    a = InStr(importNameVar, ".")
    If a > 0 Then
        importNameVar = Mid(importNameVar, a + 1)
    End If
    
    a = InStr(fileLine, "[")
    b = InStr(fileLine, "]")
    If a > 0 And b > a Then
        a = a + 1
        addrVar = Mid(fileLine, a, b - a)
    End If
    
    If KeyExistsInCollection(unique, CStr(importNameVar)) Then
        importNameVar = Empty
        addrVar = Empty
    End If
    
End Sub

Sub ImportStyleCall(fileLine, addrVar, importNameVar)
    '00402330   CALL 13_5k.00403F66                       urlmon.URLDownloadToFileA
    '                      --------                       -------------------------
    On Error Resume Next
    
    l = Split(Trim(fileLine), " ") 'words (we want first(address) and last (api name)
    importNameVar = l(UBound(l))
    
    a = InStrRev(importNameVar, ".")
    If a > 0 Then
        importNameVar = Mid(importNameVar, a + 1)
    End If
    
    a = InStr(fileLine, ".")
    If a > 25 Then 'module name not call x.
        a = 0
        b = InStr(1, fileLine, "CALL ")
    Else
        b = InStr(a, fileLine, " ")
    End If
    
    If a > 0 And b > a Then
        a = a + 1
        addrVar = Mid(fileLine, a, b - a)
    ElseIf a < 1 And b > 0 Then
        b = b + 6
        addrVar = Mid(fileLine, b, InStr((b + 1), fileLine, " ") - b)
        addrVar = Replace(addrVar, vbTab, "")
    Else
        addrVar = ""
        importNameVar = ""
        Exit Sub
    End If
        
    If KeyExistsInCollection(unique, CStr(importNameVar)) Then
        importNameVar = Empty
        addrVar = Empty
    End If
    
End Sub




Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = c(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function
    





Private Sub Text2_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Not FileExists(Data.Files(1)) Then
        MsgBox "Files only"
        Exit Sub
    End If

    f = Data.Files(1)
    Text2 = ReadFile(f)
    loadedFile = f
    
End Sub



Function ReadFile(filename)
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function



Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Private Sub Command3_Click()
    
    If Len(Text1) = 0 Then
        MsgBox "Enter expression to match, uses VB LIKE keyword", vbInformation
        Exit Sub
    End If
    
    tmp = Split(Text2, vbCrLf)
    For i = 0 To UBound(tmp)
        If tmp(i) Like Text1 Then tmp(i) = ""
    Next
    
    tmp = Join(tmp, vbCrLf)
    tmp = Replace(tmp, vbCrLf & vbCrLf, vbCrLf)
    Text2 = tmp
    
    
    
End Sub

Private Sub Text2_DblClick()
    Text2 = Clipboard.GetText
End Sub



Sub writeFile(path, it)
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub
