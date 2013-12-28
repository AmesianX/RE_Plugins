VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Active IDA Servers"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  
    Dim x
    For Each x In Servers 'remove any that arent still valid..
        If IsWindow(x) = 0 Then
            Servers.Remove "hwnd:" & x
        Else
            List1.AddItem "hwnd: " & x & " --> " & SendCmdRecvText("loadedfile:" & Form1.hwnd, x)
        End If
    Next
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Width = Me.Width - List1.Left - 200
    List1.Height = Me.Height - List1.Top - 200
End Sub
