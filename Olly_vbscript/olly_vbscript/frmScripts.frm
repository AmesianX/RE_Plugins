VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmScripts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saved Scripts"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Load"
      Height          =   375
      Left            =   3300
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin MSComctlLib.ListView lv 
      Height          =   1935
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3413
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Saved Scripts by Name"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmScripts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'dzzie@yahoo.com
'http://sandsprite.com

Dim selLi As ListItem


Private Sub cmdDelete_Click()

    On Error Resume Next
       
    If selLi Is Nothing Then
        MsgBox "No selection made", vbInformation
        Exit Sub
    End If
    
    cn.Execute "Delete from tblScripts where autoid=" & selLi.Tag
    
    lv.ListItems.Remove selLi.index
    Set selLi = Nothing

End Sub

Private Sub cmdSelect_Click()
    
    On Error Resume Next
    
    Dim rs As Recordset
    
    If selLi Is Nothing Then
        MsgBox "No selection made", vbInformation
        Exit Sub
    End If
    
    Set rs = cn.Execute("Select * from tblScripts where autoid=" & selLi.Tag)
    
    If rs.EOF Then
        MsgBox "Record not found?"
        Exit Sub
    End If
    
    frmInstance.txtScript.Text = rs!scripttext
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Dim rs As Recordset
    Dim li As ListItem
    
    lv.ColumnHeaders(1).Width = lv.Width - 15
    
    Set rs = cn.Execute("Select * from tblScripts")
    
    While Not rs.EOF
        Set li = lv.ListItems.Add
        li.Text = rs!scriptname
        li.Tag = rs!autoid
        rs.MoveNext
    Wend
    
    On Error Resume Next
    Set selLi = lv.ListItems(1)
    
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selLi = Item
End Sub
