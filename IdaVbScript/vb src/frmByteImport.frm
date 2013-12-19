VERSION 5.00
Begin VB.Form frmByteImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patch in Bytes from File to offset"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   7380
      TabIndex        =   9
      Top             =   60
      Width           =   675
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do it "
      Height          =   375
      Left            =   6780
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtEA 
      Height          =   315
      Left            =   5520
      TabIndex        =   7
      Top             =   480
      Width           =   1035
   End
   Begin VB.TextBox txtFileReadLen 
      Height          =   315
      Left            =   3300
      TabIndex        =   5
      Top             =   480
      Width           =   1035
   End
   Begin VB.TextBox txtStart 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   915
   End
   Begin VB.TextBox txtFile 
      Height          =   315
      Left            =   1740
      Locked          =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Text            =   "Drag & Drop file here"
      Top             =   60
      Width           =   5475
   End
   Begin VB.Label Label1 
      Caption         =   "IDB EA Offset"
      Height          =   195
      Index           =   3
      Left            =   4440
      TabIndex        =   6
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "File Read Len:"
      Height          =   195
      Index           =   2
      Left            =   2160
      TabIndex        =   4
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Start File Offset:"
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "File to Patch into IDB"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmByteImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private myFile As String

 

Private Sub cmdBrowse_Click()
    Dim c As New clsCmnDlg
    Dim s As String
    
    s = c.OpenDialog(AllFiles)
    If Len(s) = 0 Then Exit Sub
    
    Process s
    
End Sub

Private Sub Command1_Click()
    On Error GoTo hell
    
    If Len(myFile) = 0 Or Not FileExists(myFile) Then
        MsgBox "File not found"
        Exit Sub
    End If
    
    Dim start As Long, leng As Long, ea As Long
    
    start = CLng("&h" & txtStart)
    leng = CLng("&h" & txtFileReadLen)
    ea = CLng("&h" & txtEA)
    
    Dim b() As Byte
    Dim f As Long
    
    leng = leng - 1
    
    ReDim b(leng)
    f = FreeFile
    
    Open myFile For Binary As f
    Get f, start + 1, b()
    Close f
    
    Dim i As Long
    
    For i = 0 To leng
        PatchByte ea, b(i)
        ea = ea + 1
    Next
    
    MsgBox "Patching Complete", vbInformation
    
    frmPluginSample.Done
    
    Exit Sub
hell: MsgBox Err.Description
End Sub

Private Sub txtFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim f As String
    f = Data.Files(1)
    Process f
End Sub

Sub Process(f)
    If FileExists(f) Then
        myFile = f
        txtStart = 1
        txtFileReadLen = Hex(FileLen(f))
        txtFile = f
    Else
        myFile = Empty
    End If
End Sub
