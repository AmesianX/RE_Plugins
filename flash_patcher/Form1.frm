VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{9A143468-B450-48DD-930D-925078198E4D}#1.0#0"; "hexed.ocx"
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "FlashPatcher"
   ClientHeight    =   10335
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   23070
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   23070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpenSelection 
      Caption         =   "Selection To New File"
      Height          =   375
      Left            =   20820
      TabIndex        =   15
      Top             =   60
      Width           =   2055
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   375
      Left            =   7920
      TabIndex        =   14
      Top             =   60
      Width           =   735
   End
   Begin VB.CommandButton cmdSaveEdits 
      Caption         =   "Save Patches"
      Height          =   375
      Left            =   19380
      TabIndex        =   12
      Top             =   60
      Width           =   1215
   End
   Begin VB.Frame fraOpts 
      Caption         =   "Options"
      Height          =   8355
      Left            =   660
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   14475
      Begin VB.ListBox List1 
         Height          =   7080
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   14055
      End
      Begin VB.TextBox txtFlexBin 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         TabIndex        =   7
         Top             =   240
         Width           =   6975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   13980
         TabIndex        =   10
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Debug Log (Double click an entry to view msgbox)"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   720
         Width           =   4635
      End
      Begin VB.Label Label2 
         Caption         =   "Flex BinDir"
         Height          =   375
         Left            =   60
         TabIndex        =   8
         Top             =   300
         Width           =   915
      End
   End
   Begin rhexed.HexEd he 
      Height          =   8835
      Left            =   11400
      TabIndex        =   5
      Top             =   480
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   15584
   End
   Begin VB.CommandButton cmdDecompile 
      Caption         =   "Decompile"
      Height          =   375
      Left            =   8700
      TabIndex        =   4
      Top             =   60
      Width           =   1275
   End
   Begin VB.TextBox txtswf 
      Height          =   315
      Left            =   840
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Text            =   "c:\bad.swf"
      Top             =   60
      Width           =   6975
   End
   Begin MSComctlLib.ListView lvFunc 
      Height          =   5235
      Left            =   120
      TabIndex        =   1
      Top             =   4020
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   9234
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Opcodes"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Disasm"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Function"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblLoadedFile 
      Height          =   315
      Left            =   11520
      TabIndex        =   13
      Top             =   60
      Width           =   7755
   End
   Begin VB.Label Label1 
      Caption         =   "SWF File"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuStringPool 
         Caption         =   "Show String Pool"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowRawDisasm 
         Caption         =   "Show Raw Disasm"
      End
      Begin VB.Menu mnuDelCachedDisasm 
         Caption         =   "Delete Cached Disasm"
      End
      Begin VB.Menu mnuDeleteCachedDecompressed 
         Caption         =   "Delete Cached Decompressed SWF"
      End
      Begin VB.Menu mnuSpacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuASRef 
         Caption         =   "ActionScript Reference"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Debug Log"
      End
   End
   Begin VB.Menu mnuFunc 
      Caption         =   "mnuFunc"
      Begin VB.Menu mnuNopSelection 
         Caption         =   "Nop Selected"
      End
      Begin VB.Menu mnuOriginalBytes 
         Caption         =   "Original Bytes"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchFor 
         Caption         =   "Search For"
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find Next"
      End
      Begin VB.Menu mnuGotoOffset 
         Caption         =   "Goto Offset"
      End
   End
   Begin VB.Menu mnuLv 
      Caption         =   "mnuLv"
      Begin VB.Menu mnuCopyDisasm 
         Caption         =   "Copy Disasm"
      End
      Begin VB.Menu mnuHtmlView 
         Caption         =   "Html Viewer"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim li As ListItem
Dim selFuncLi As ListItem
Dim selLi As ListItem
Dim ActiveFunction As CFunction

Private Sub cmdBrowse_Click()
    Dim x As String
    x = dlg.OpenDialog(AllFiles, , , Me.hWnd)
    If Len(x) = 0 Then Exit Sub
    txtswf = x
    cmdDecompile_Click
End Sub

Private Sub cmdDecompile_Click()
        
    lv.ListItems.Clear
    lvFunc.ListItems.Clear
    
    If Not FileExists(txtswf) Then
        MsgBox "File not found!", vbInformation
        Exit Sub
    End If
    
    If Not p.Decompile(txtswf) Then
        MsgBox "decompilation failed see debug log on cfg pane"
        Exit Sub
    End If
    
    If Not p.LoadDecompilation() Then
        MsgBox "Loading Decomp Failed"
        Exit Sub
    End If
    
    lblLoadedFile.Caption = "Showing: " & p.DecompressedSWF
    he.LoadFile p.DecompressedSWF, False
    
    Dim f As CFunction
    Dim bd As CBinaryData
    
    For Each bd In p.BinaryData
        Set li = lv.ListItems.Add(, , Hex(bd.Offset))
        li.SubItems(1) = bd.Size
        li.SubItems(2) = "DefineBinaryData ID: " & bd.ID
        Set li.Tag = bd
    Next
    
    For Each f In p.Functions
        Set li = lv.ListItems.Add(, , Hex(f.StartOffset))
        li.SubItems(1) = f.CodeLength
        li.SubItems(2) = f.Prototype
        Set li.Tag = f
    Next
    
End Sub

 

Private Sub cmdOpenSelection_Click()
    Dim f As CHexEditor
    Dim b() As Byte
    Dim ff As Long
    
    On Error GoTo hell
    
    If he.SelLength = 0 Then Exit Sub
    
    ReDim b(he.SelLength - 1)
    
    ff = FreeFile
    Open p.DecompressedSWF For Binary As ff
    Get ff, he.SelStart + 1, b()
    Close ff
    
    Set f = New CHexEditor
    f.Editor.LoadByteArray b(), False
    
hell:
    
End Sub

Private Sub cmdSaveEdits_Click()
    he.Save
End Sub

Private Sub Form_Load()

    
    
    LV_LastColumnResize lv
    LV_LastColumnResize lvFunc
    
    'lv.ColumnHeaders(lv.ColumnHeaders.Count).Width = lv.Width - lv.ColumnHeaders(lv.ColumnHeaders.Count).Left - 100
    'lvFunc.ColumnHeaders(lvFunc.ColumnHeaders.Count).Width = lvFunc.Width - lvFunc.ColumnHeaders(lvFunc.ColumnHeaders.Count).Left - 100
    Set p.dbg = List1
    mnuFunc.Visible = False
    mnuLv.Visible = False
    
    'txtFlexBin = GetSetting("flashPatcher", "settings", "FlexBin", "D:\_Lilguys\flash tools\flex_sdk_4.1.0.16076\bin")
    'txtFlexBin = GetSetting("flashPatcher", "settings", "FlexBin", App.Path & "\swfdump4.1\bin")
    txtFlexBin = App.path & "\swfdump4.1\bin" 'now included...
    txtswf = GetSetting("flashPatcher", "settings", "lastSwf", "c:\bad.swf")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
     SaveSetting "flashPatcher", "settings", "FlexBin", txtFlexBin
     SaveSetting "flashPatcher", "settings", "lastSwf", txtswf
End Sub

Private Sub Label4_Click()
    fraOpts.Visible = False
End Sub

Private Sub List1_DblClick()
    MsgBox List1.List(List1.ListIndex)
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim f As CFunction
    Dim obj As Object
    Dim bd As CBinaryData
    
    Set selLi = Item
    Set obj = Item.Tag
    
    If TypeName(obj) = "CBinaryData" Then
        Set bd = Item.Tag
        he.scrollTo bd.Offset
        he.SelStart = bd.Offset
        he.SelLength = bd.Size
        Exit Sub
    End If
        
    Set f = Item.Tag
    
    If ObjPtr(ActiveFunction) = ObjPtr(f) Then Exit Sub 'already displaying dont reload
    
    Set ActiveFunction = f
    
    Dim ci As CInstruction
    
    lvFunc.ListItems.Clear
    
    Dim o As Long
    o = f.StartOffset
    
    he.scrollTo o
    he.SelStart = o
    he.SelLength = f.CodeLength
    
    For Each ci In f.Instructions
        Set li = lvFunc.ListItems.Add(, , Hex(ci.Offset))
        li.SubItems(1) = ci.OpCodes
        li.SubItems(2) = ci.Disasm
        If ci.isLabel Then SetLiColor li, &H808000 Else If ci.isPossibleBranch Then SetLiColor li, vbBlue
        Set li.Tag = ci
    Next
        
    
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuLv
End Sub

Private Sub lvFunc_DblClick()
    If selFuncLi Is Nothing Then Exit Sub
    Dim ci As CInstruction
    Dim ci2 As CInstruction
    Dim tmp As String
    
    Set ci = selFuncLi.Tag
    
    a = InStrRev(ci.Disasm, "L")
    If a > 0 Then 'possible branch label
        tmp = Trim(Mid(ci.Disasm, a))
        Set ci2 = ActiveFunction.FindLabel(tmp)
        If Not ci2 Is Nothing Then
            'selFuncLi.Selected = False
            JumpToCI ci2
        End If
    End If
        
End Sub

Function JumpToCI(ci As CInstruction)
    Dim x As CInstruction
    
    For Each li In lvFunc.ListItems
        Set x = li.Tag
        If ObjPtr(x) = ObjPtr(ci) Then
            li.EnsureVisible
            li.Selected = True
            Exit Function
        End If
    Next
    
End Function

Public Sub lvFunc_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Set selFuncLi = Item
    
    Dim ci As CInstruction
    Set ci = Item.Tag
    
    Dim o As Long
    o = CLng("&h" & Item.Text)
    
    he.scrollTo o
    he.SelStart = o
    he.SelLength = ci.InstructionLength
    
End Sub

Private Sub lvFunc_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuFunc
End Sub

Private Sub txtOut_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    txtOut = Data.Files(1)
End Sub

Private Sub mnuAbout_Click()
    MsgBox "FlashPatcher v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & _
            "UI developer David Zimmer <dzzie@yahoo.com>" & vbCrLf & _
            "swfdump/FlexSDK copyright Adobe" & vbCrLf & _
            "hexed.ocx copyright Rang3r" & vbCrLf & _
            "zlib.dll  copyright Jean-loup Gailly and Mark Adler" & vbCrLf & _
            "cmdoutput.bas copyright Joacim Andersson, Brixoft Software" _
            , vbInformation
End Sub

Private Sub mnuASRef_Click()
    On Error Resume Next
    Shell "cmd /c start http://learn.adobe.com/wiki/display/AVM2/add", vbHide
End Sub

Private Sub mnuCopyDisasm_Click()
    
    If ActiveFunction Is Nothing Then Exit Sub
    
    Dim tmp As String
    Dim ci As CInstruction
    Dim vt As CVariable
    Dim ret()
    
    For Each vt In ActiveFunction.varTypes
        push ret, " " & vt.varName & " : " & vt.varType & vbCrLf
    Next
    
    push ret, vbCrLf
    
    For Each ci In ActiveFunction.Instructions
        If ci.isLabel Then tmp = tmp & vbCrLf & pad(Hex(ci.Offset)) & pad(" ", 12) & ci.Label & ":" & vbCrLf
        d = ci.Disasm
        If ci.isLabel Then
            If InStr(d, "label") > 0 Then GoTo nextone
            d = Trim(Replace(d, ci.Label & ":", Empty))
        End If
        
    
        push ret, " " & pad(Hex(ci.Offset)) & pad(ci.OpCodes, 16) & d & vbCrLf
nextone:
    Next
    
    Clipboard.Clear
    Clipboard.SetText Join(ret, "")
    
End Sub



Private Sub mnuDelCachedDisasm_Click()
    If FileExists(p.CurrentDecompilation) Then Kill p.CurrentDecompilation
    lv.ListItems.Clear
    lvFunc.ListItems.Clear
End Sub

Private Sub mnuDeleteCachedDecompressed_Click()
    If p.CurSWF = p.DecompressedSWF Then
        If MsgBox("Current Decompressed SWF is the main SWF file you loaded are you sure you wish to delete it?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    If FileExists(p.DecompressedSWF) Then Kill p.DecompressedSWF
    he.LoadString ""
    mnuDelCachedDisasm_Click
End Sub

Private Sub mnuFindNext_Click()
    If selFuncLi Is Nothing Then Exit Sub
    
    x = InputBox("Enter command to search for:")
    If Len(x) = 0 Then Exit Sub
    
    Dim startFound As Boolean
    Dim ci As CInstruction
    
    For Each li In lvFunc.ListItems
    
        If startFound Then
            Set ci = li.Tag
            If InStr(1, ci.Disasm, x, vbTextCompare) > 0 Then
                li.EnsureVisible
                li.Selected = True
                Exit For
            End If
        End If
        
        If ObjPtr(li) = ObjPtr(selFuncLi) Then
            startFound = True
        End If
        
    Next
    
End Sub

Private Sub mnuGotoOffset_Click()
    
    Dim ci As CInstruction
    Dim Offset As Long
    On Error Resume Next
    
    x = InputBox("Enter HEX offset to jump to")
    If Len(x) = 0 Then Exit Sub
    
    Offset = CLng("&h" & x)
    If Err.Number <> 0 Then
        MsgBox "Invalid HexNumber"
        Exit Sub
    End If
    
    If Offset = 0 Then Exit Sub
    
    For Each li In lvFunc.ListItems
        Set ci = li.Tag
        If ci.Offset = Offset Then
            li.EnsureVisible
            li.Selected = True
            Exit Sub
        End If
    Next
    
    MsgBox "Offset " & Hex(Offset) & " not found"
    
End Sub

Private Sub mnuHtmlView_Click()

    If ActiveFunction Is Nothing Then Exit Sub
    
    Dim tmp As String
    Dim ci As CInstruction
    Dim vt As CVariable
    
    Dim ret()
    
    For Each vt In ActiveFunction.globals
        h = "global <u><A href=""javascript:redo('g_','XXXX')"" style='color:#006699;cursor:pointer' NAME='#XXXX' onclick=""return -1"">XXXX</A></u> : " & vt.varType
        h = Replace(h, "XXXX", vt.varName) & vbCrLf
        push ret, h
    Next
    
    push ret, vbCrLf
    
    For Each vt In ActiveFunction.varTypes
        h = "<u><A href=""javascript:redo('var_','XXXX')"" style='color:#006699;cursor:pointer' NAME='#XXXX' onclick=""return -1"">XXXX</A></u> : " & vt.varType
        h = Replace(h, "XXXX", vt.varName) & vbCrLf
        push ret, h
    Next
    
    push ret, vbCrLf
    
    
    For Each ci In ActiveFunction.Instructions
        
        tmp = Empty
        
        If ci.isLabel Then
            h = "<u><A href=""javascript:redo('lbl_','XXXX')"" style='color:#006699;cursor:pointer' NAME='#XXXX' onclick=""return -1"">XXXX</A></u> "
            h = Replace(h, "XXXX", ci.Label & "_")
            tmp = tmp & vbCrLf & pad(Hex(ci.Offset)) & pad(" ", 12) & h & ":" & vbCrLf
        End If
        
        d = ci.Disasm
        
        If ci.isLabel Then
            If InStr(d, "label") > 0 Then GoTo nextone
            d = Trim(Replace(d, ci.Label & ":", Empty))
        End If
        
        If ci.isPossibleBranch Then
            link = "<a href='#XXXX'>XXXX</a> "
            link = Replace(link, "XXXX", ci.branchTarget & "_")
            d = Replace(d, ci.branchTarget, link)
        End If
        
        If Len(ci.variableName) > 0 Then
            h = "<u><A href=""javascript:redo('var_','XXXX')"" style='color:#006699;cursor:pointer' NAME='#XXXX' onclick=""return -1"">XXXX</A></u> "
            h = Replace(h, "XXXX", ci.variableName)
            d = Replace(d, ci.variableName, h)
        End If
        
        'span with &nbsp works the way i want, but its slow to load or rename labels..
        d = d & "<font color='#009900'><span contenteditable='true'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></font>"
        'd = d & "<div contenteditable='true'><font color='#009900'>                                   </font></div>"
        
        'd = d & "<font color='#009900'>                                                                                          </font>"
        
        push ret(), tmp & pad(Hex(ci.Offset)) & pad(ci.OpCodes, 16) & d & vbCrLf
nextone:
    Next
    
    frmBrowser.LoadGraph Join(ret, "")
    
End Sub

Private Sub mnuNopSelection_Click()
    Dim li As ListItem
    Dim ci As CInstruction
    
    Const flashNop = 2
    Dim b() As Byte
    
    For Each li In lvFunc.ListItems
        If li.Selected Then NopLi li
    Next
    
End Sub

Public Sub NopLi(li As ListItem)
    Dim ci As CInstruction
    Const flashNop = 2
    Dim b() As Byte
    
    li.EnsureVisible
    Set ci = li.Tag
    ReDim b(ci.InstructionLength - 1)
    For i = 0 To ci.InstructionLength - 1
        b(i) = flashNop
    Next
    Form1.he.OverWriteData CLng("&h" & li.Text), b()
    li.SubItems(2) = "NOPPED   was -> " & li.SubItems(2)
    SetLiColor li, vbRed
    li.Selected = False
            
End Sub

Private Sub mnuOptions_Click()
    fraOpts.Visible = True
End Sub

Private Sub mnuOriginalBytes_Click()
    Dim li As ListItem
    Dim ci As CInstruction
    
    Dim b() As Byte
    
    For Each li In lvFunc.ListItems
        If li.Selected Then
            li.EnsureVisible
            Set ci = li.Tag
            b() = StringOpcodesToBytes(ci.OpCodes)
            Form1.he.OverWriteData CLng("&h" & li.Text), b()
            li.SubItems(2) = ci.Disasm
            SetLiColor li, vbBlack
            li.Selected = False
        End If
    Next
    
    
End Sub

Private Sub mnuSearchFor_Click()
    Dim x As String
    x = InputBox("Enter disasm to search for:")
    If Len(x) = 0 Then Exit Sub
    frmSearch.SearchFor x
End Sub

Private Sub mnuShowRawDisasm_Click()
    On Error Resume Next
    Shell "notepad.exe " & p.CurrentDecompilation, vbNormalFocus
End Sub

Private Sub mnuStringPool_Click()
    frmStringPool.Visible = True
End Sub

Private Sub txtswf_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    txtswf = Data.Files(1)
End Sub
