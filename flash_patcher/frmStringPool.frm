VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStringPool 
   Caption         =   "String Pool Viewer"
   ClientHeight    =   7485
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6960
   LinkTopic       =   "Form2"
   ScaleHeight     =   7485
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lv 
      Height          =   7395
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   13044
      View            =   3
      LabelEdit       =   1
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
         Text            =   "Index"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Offset"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "String"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuFindReferences 
         Caption         =   "Find References"
      End
   End
End
Attribute VB_Name = "frmStringPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selspi As CStringPoolItem

Private Sub Form_Load()
    
    Dim spi As CStringPoolItem
    Dim li As ListItem
    
    mnuPopup.Visible = False
    
    For Each spi In p.StringPool
        Set li = lv.ListItems.Add(, , Hex(spi.Index))
        li.SubItems(1) = spi.RawOffset & " / " & Hex(spi.Offset)
        li.SubItems(2) = spi.Data
        Set li.Tag = spi
    Next
                
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lv.Width = Me.Width - lv.Left - 200
    lv.Height = Me.Height - lv.Top - 450
    LV_LastColumnResize lv
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim spi As CStringPoolItem
    Set spi = Item.Tag
    Set selspi = spi
    Form1.he.scrollTo spi.Offset
    Form1.he.SelStart = spi.Offset
    Form1.he.SelLength = spi.DataLength + 1
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuFindReferences_Click()
    
    If selspi Is Nothing Then Exit Sub
    
    Dim f As CFunction
    Dim ci As CInstruction
    Dim matches As New Collection
    Dim tmp
    
    For Each f In p.Functions
        For Each ci In f.Instructions
            If InStr(ci.Disasm, selspi.Data) > 0 Then
                matches.Add "Function: " & f.Prototype & vbCrLf & vbTab & _
                            "Offset: " & Hex(ci.Offset) & vbCrLf & vbTab & _
                            "Disasm: " & ci.Disasm
            End If
        Next
    Next
    
    For Each x In matches
        tmp = tmp & x & vbCrLf & vbCrLf
    Next
    
    Dim ff As String
    ff = GetFreeFileName(Environ("temp"))
    WriteFile ff, tmp
    ff = GetShortName(ff)
    
    On Error Resume Next
    Shell "notepad.exe " & ff, vbNormalFocus
    
    
End Sub
