VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLoadScript 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   315
      Left            =   2100
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      Height          =   315
      Left            =   3420
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Script Name"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmLoadScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SelLi As ListItem


Private Sub cmdDelete_Click()
    
     If SelLi Is Nothing Then Exit Sub
     
     If MsgBox("Are you sure you want to delete script: " & SelLi.Text, vbYesNo) = vbYes Then
            frmIDAScript.cn.Execute "Delete from tblScripts where autoid=" & SelLi.Tag
            lv.ListItems.Remove SelLi.index
            Set SelLi = Nothing
            MsgBox "Script Deleted", vbInformation
     End If
     
End Sub

Private Sub Command1_Click()

    If SelLi Is Nothing Then Exit Sub
    
    Dim txt As String
    On Error GoTo hell
    
    txt = frmIDAScript.cn.Execute("Select scriptText from tblScripts where autoid=" & SelLi.Tag)!scripttext
    frmIDAScript.txtScript.Text = txt
    frmIDAScript.selectedID = CLng(SelLi.Tag)
    frmIDAScript.selectedName = SelLi.Text
    
hell:
    Unload Me
End Sub

Private Sub Form_Load()

    Dim rs As Recordset
    Dim li As ListItem
    
    lv.ColumnHeaders(1).Width = lv.Width - 70
    
    Set rs = frmIDAScript.cn.Execute("Select * from tblScripts")
    
    If rs.EOF Then
        'MsgBox "No existing items sorry", vbInformation
        Exit Sub
    End If
    
    While Not rs.EOF
        Set li = lv.ListItems.Add
        li.Text = rs!scriptName
        li.Tag = rs!autoid
        rs.MoveNext
    Wend
    
    Set SelLi = lv.ListItems(1)
    
    

End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set SelLi = Item
End Sub


