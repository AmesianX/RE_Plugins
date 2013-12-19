VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSearch 
   Caption         =   "Search"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11310
   LinkTopic       =   "Form2"
   ScaleHeight     =   5865
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select All"
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   5340
      Width           =   1695
   End
   Begin VB.CommandButton cmdRemvSel 
      Caption         =   "Remove Selected"
      Height          =   435
      Left            =   3120
      TabIndex        =   3
      Top             =   5340
      Width           =   1755
   End
   Begin VB.CommandButton cmdNopAll 
      Caption         =   "Nop All"
      Height          =   495
      Left            =   9120
      TabIndex        =   2
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton cmdNopSelected 
      Caption         =   "Nop Selected"
      Height          =   495
      Left            =   6300
      TabIndex        =   1
      Top             =   5280
      Width           =   1935
   End
   Begin MSComctlLib.ListView lvFunc 
      Height          =   5235
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Function SearchFor(x As String)
    Me.Visible = True
    
    Dim li As ListItem
    Dim li2 As ListItem
    
    lvFunc.ListItems.Clear
    
    For Each li In Form1.lvFunc.ListItems
        If InStr(1, li.SubItems(2), x, vbTextCompare) > 0 Or InStr(1, li.SubItems(1), x, vbTextCompare) > 0 Then
            Set li2 = lvFunc.ListItems.Add(, , li.Text)
            li2.SubItems(1) = li.SubItems(1)
            li2.SubItems(2) = li.SubItems(2)
            Set li2.Tag = li
        End If
    Next
    
    
    
End Function

Private Sub cmdNopAll_Click()

    Dim li As ListItem
    Dim form1_li As ListItem
    
    For Each li In Me.lvFunc.ListItems
        Set form1_li = li.Tag
        Form1.NopLi form1_li
        li.SubItems(2) = "NOPPED   was -> " & li.SubItems(2)
    Next
    
End Sub

Private Sub cmdNopSelected_Click()
    
    Dim li As ListItem
    Dim form1_li As ListItem
    
    For Each li In Me.lvFunc.ListItems
        If li.Selected Then
            Set form1_li = li.Tag
            Form1.NopLi form1_li
            li.SubItems(2) = "NOPPED   was -> " & li.SubItems(2)
        End If
    Next
    
End Sub

Private Sub cmdRemvSel_Click()
    For i = lvFunc.ListItems.Count To 1 Step -1
        If lvFunc.ListItems(i).Selected Then lvFunc.ListItems.Remove i
    Next
End Sub

Private Sub cmdSelectAll_Click()
    Dim li As ListItem
    Dim li2 As ListItem
    
    For Each li In lvFunc.ListItems
         Set li2 = li.Tag 'lvFunc listitem object...
         li2.EnsureVisible
         li2.Selected = True
    Next
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    'lvFunc.Width = Me.Width - lvFunc.Left - 200
    'lvFunc.Height = Me.Height - lvFunc.Top - 200
    LV_LastColumnResize lvFunc
End Sub


Private Sub lvFunc_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim li As ListItem
    Set li = Item.Tag
    li.EnsureVisible
    li.Selected = True
    Form1.lvFunc_ItemClick li
End Sub
