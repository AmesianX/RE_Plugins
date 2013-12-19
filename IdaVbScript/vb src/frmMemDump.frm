VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMemDump 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6690
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDump 
      Caption         =   "Dump"
      Height          =   315
      Left            =   5340
      TabIndex        =   5
      Top             =   1380
      Width           =   1275
   End
   Begin VB.TextBox txtLen 
      Height          =   315
      Left            =   3840
      TabIndex        =   4
      Top             =   1380
      Width           =   1215
   End
   Begin VB.TextBox txtStart 
      Height          =   315
      Left            =   1860
      TabIndex        =   2
      Top             =   1380
      Width           =   1275
   End
   Begin MSComctlLib.ListView lvSect 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2355
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Virtual Addr"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Virtual Size"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "RawOffset"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "RawSize"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Attributes"
         Object.Width           =   1235
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Start                                      Len"
      Height          =   255
      Index           =   1
      Left            =   1380
      TabIndex        =   3
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Dump Memory :"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   1440
      Width           =   1155
   End
End
Attribute VB_Name = "frmMemDump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 

Private Sub Form_Load()
    
    Dim curFile As String
    
    curFile = loadedFile
    
    If curFile <> clsSect.curFile Then
        clsSect.LoadSections curFile
    End If
    
    clsSect.FilloutListView lvSect
    Me.Caption = "ImageBase: " & Hex(clsSect.ImageBase)

End Sub

Private Sub lvSect_ItemClick(ByVal Item As MSComctlLib.ListItem)
     
     txtStart = Item.SubItems(1)
     txtLen = Item.SubItems(2)
             
     Dim s As Long
     GetHextxt txtStart, s
     
     s = s + clsSect.ImageBase
     txtStart = Hex(s)
     
End Sub


Private Sub cmdDump_Click()

    On Error GoTo hell
        
    Dim s  As Long, l As Long
    
    If Not GetHextxt(txtStart, s) Then Exit Sub
    If Not GetHextxt(txtLen, l) Then Exit Sub
    
    If s < clsSect.ImageBase Then
        MsgBox "Start less than Imagebase?"
        Exit Sub
    End If
    
    If l < 1 Then
        MsgBox "Can cannot be 0 or negative"
        Exit Sub
    End If
    
    Dim dlg As New clsCmnDlg
    Dim fname As String
    Dim f As Long
    

    fname = dlg.SaveDialog(AllFiles, , "Save Dump as")
    If Len(fname) = 0 Then Exit Sub
    
    If FileExists(fname) Then Kill fname
        
    Dim buf() As Byte
    ReDim buf(1 To l)
    
    GetBytes s, buf(1), l
    
    f = FreeFile
    Open fname For Binary As f
    Put f, , buf()
    Close f
    
    MsgBox "Dump Complete", vbInformation
    
    
 Exit Sub
hell:  MsgBox Err.Description
    
End Sub




