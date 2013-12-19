VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTextSearch 
   Caption         =   "Text Search - Add Multi Comments"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5520
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Stop"
      Height          =   375
      Left            =   9600
      TabIndex        =   14
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "cpy"
      Height          =   375
      Left            =   9720
      TabIndex        =   13
      Top             =   6840
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "clr"
      Height          =   375
      Left            =   9720
      TabIndex        =   12
      Top             =   7200
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   11
      Top             =   6360
      Width           =   9615
   End
   Begin MSScriptControlCtl.ScriptControl sc 
      Left            =   9720
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      Timeout         =   1000000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Run"
      Height          =   375
      Left            =   8400
      TabIndex        =   10
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   375
      Left            =   7560
      TabIndex        =   9
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Top             =   5880
      Width           =   5175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Comment"
      Height          =   375
      Left            =   8880
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4935
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Offset"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Disasm"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Comment"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Run Vbs Automation Script"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Set Comment for selected"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Search Text"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmTextSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmndlg As New clsCmnDlg
Dim fso As New clsFileSystem

Function FunctionAtVA(va As Long) As CFunction
    Set FunctionAtVA = Module1.FunctionAtVA(va)
End Function

Function DoSearch(Optional parameter = "") As Long
        
        If Len(parameter) > 0 Then
            Text1 = parameter
        End If
        
        Command1_Click
        
        DoSearch = lv.ListItems.count
        
End Function

Function AddComments(Optional comment = "")

    If Len(comment) > 0 Then
        Text2 = comment
    End If
    
    Command2_Click

End Function


Sub Command1_Click()
    
    Dim x, tmp
    Dim addr As Long
    Dim laddr As Long
    Dim Search As Searchtype
    Dim li As ListItem
    
    x = Text1
    tmp = ""
    
    addr = 1
    Search = SEARCH_DOWN
    lv.ListItems.Clear
    
    Do While 1 'test for addr > 0 fails because addr is signed duh!
        laddr = addr
        'MsgBox Hex(addr)
        addr = SearchText(addr, x, Search, 0)
        If addr = -1 Then Exit Do
        tmp = tmp & Hex(addr) & " " & GetAsmCode(addr) & vbCrLf
        Set li = lv.ListItems.Add(, , Hex(addr))
        li.SubItems(1) = GetAsmCode(addr)
        li.SubItems(2) = Get_Comment(addr)
        addr = addr + InstructionLength(addr)
        'MsgBox "2: new addr " & Hex(addr) & " last addr: " & Hex(laddr)
        'If laddr = addr Then Exit Do
        DoEvents
    Loop

    DoEvents
    Me.Refresh
    
    'List1.AddItem lv.ListItems.Count & " matches found!"
    
End Sub

Sub Command2_Click()
    On Error GoTo hell
    
    Dim li As ListItem
    Dim a As Long
    Dim changed As Long
    Dim txt As String
    Dim ret As Long
    
    For Each li In lv.ListItems
        If Not li Is Nothing And Len(li.Text) > 0 Then
1           a = CLng("&h" & li.Text)
2           r = Module1.Set_Comment(a, Text2.Text)
3           If r = 0 Then List1.AddItem a & " " & Hex(a) & " " & Text2.Text & " ret=" & r
            changed = changed + 1
        End If
    Next
        
    Me.Refresh
    DoEvents
    
    If changed = 0 Then
        List1.AddItem "No entries selected?", vbInformation
    Else
        List1.AddItem changed & " comments added!", vbInformation
    End If
    
Exit Sub
        
        
hell:
        If Not li Is Nothing Then lit = li.Text Else lit = "[li is nothing]"
       ' List1.AddItem "Error in Command2_Click: Line:" & Erl() & " " & Err.Description & lit & " " & Text2.Text
        Resume Next
        
End Sub

Function SelAll()
    On Error Resume Next
    
    Dim li As ListItem
    For Each li In lv.ListItems
        li.Selected = True
    Next
    DoEvents
    Me.Refresh
    
End Function


Private Sub Command3_Click()
    Dim f As String
    f = cmndlg.OpenDialog(AllFiles)
    If Len(f) > 0 Then Text3 = f
End Sub

Private Sub Command4_Click()
    On Error GoTo hell
    
    If fso.FileExists(Text3) Then
        t = fso.ReadFile(Text3)
    Else
        MsgBox "FIle not found: " & Text3
        Exit Sub
    End If
    
    sc.Reset
    sc.AddObject "form", Me
    sc.AddObject "txtSearch", Text1
    sc.AddObject "txtComment", Text2
    sc.AddObject "lv", lv
    sc.AddObject "pb", pb
    sc.AddObject "list1", List1
    
    sc.AddObject "fso", fso
    sc.AddObject "cmndlg", cmndlg
    sc.AddObject "clipboard", Clipboard
    sc.AddCode t
    
    If InStr(t, "main()") > 0 Then
        sc.ExecuteStatement "main()"
    End If
    
    Exit Sub
hell:     MsgBox Err.Description
    
End Sub

Private Sub Command5_Click()
    List1.Clear
End Sub

Private Sub Command6_Click()
    On Error Resume Next
    Dim x As String
    For i = 0 To List1.ListCount
        x = x & List1.List(i) & vbCrLf
    Next
    Clipboard.Clear
    Clipboard.SetText x
End Sub

Private Sub Command7_Click()
    sc.Reset
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Module1.Jump CLng("&h" & Item.Text)
End Sub

Private Sub sc_Error()
    On Error Resume Next
    List1.AddItem "Script Error line: " & sc.Error.Line & " " & sc.Error.Text
    List1.AddItem "   Description: " & sc.Error.Description
End Sub

Function GetAsmCode(offset) As String
    GetAsmCode = Module1.GetAsmCode(offset)
End Function

Function InstructionLength(offset) As Long
    InstructionLength = Module1.InstructionLength(CLng(offset))
End Function

Function Set_Comment(offset, comm) As Long
    Set_Comment = Module1.Set_Comment(offset, CStr(comm))
End Function

Function ScanForInstruction(offset, find_inst, scan_x_lines) As Long
    ScanForInstruction = Module1.ScanForInstruction(CLng(offset), CStr(find_inst), CLng(scan_x_lines))
End Function

Function AddXRef(ref_to, ref_from)
    Module1.AddCodeXRef CLng(ref_to), CLng(ref_from)
End Function

Function Setname(offset, Name)
    Module1.Setname CLng(offset), CStr(Name)
End Function


