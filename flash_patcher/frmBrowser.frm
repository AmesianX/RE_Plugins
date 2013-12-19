VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmBrowser 
   Caption         =   "Html Viewer"
   ClientHeight    =   8070
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   13620
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "Find"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   60
      Width           =   855
   End
   Begin VB.TextBox txtJS 
      Height          =   3015
      Left            =   7920
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmBrowser.frx":0000
      Top             =   180
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save As"
      Height          =   375
      Left            =   5460
      TabIndex        =   3
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Forward"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   60
      Width           =   1335
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   13560
      ExtentX         =   23918
      ExtentY         =   13150
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2670
      Top             =   2325
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":025E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0540
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0822
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":10C8
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LoadedFile As String

Function LoadGraph(html As String)
    On Error GoTo hell
    
    Style = "<head><style></style></head>" & vbCrLf
    
    html = Style & txtJS & vbCrLf & "<body con-tenteditable='true'><pre>" & vbCrLf & html & vbCrLf & "</pre></body>"
    LoadedFile = GetFreeFileName(Environ("temp"), ".html")
    WriteFile LoadedFile, html
    wb.Navigate2 "file://" & LoadedFile
    Me.Visible = True
    Exit Function
    
hell:
    MsgBox Err.Description
End Function

Private Sub Command1_Click()
    wb.GoBack
End Sub

Private Sub Command2_Click()
    wb.GoForward
End Sub

Private Sub Command3_Click()
    On Error GoTo hell
    Dim f As String, ff As String, x As String
    
    f = GetParentFolder(Form1.txtswf)
    ff = GetBaseName(Form1.txtswf) & ".disasm.html"
    x = dlg.SaveDialog(htmlFiles, f, , , Me.hWnd, ff)
    If Len(x) = 0 Then Exit Sub
    If FileExists(x) Then Kill x
    
    WriteFile x, txtJS & vbCrLf & "<pre>" & vbCrLf & wb.Document.body.innerHTML & vbCrLf & "</pre>"
    
    
    Exit Sub
hell:
    MsgBox Err.Description
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    wb.SetFocus
    SendKeys "^f"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    wb.Width = Me.Width - wb.Left - 200
    wb.Height = Me.Height - wb.Top - 500
End Sub
