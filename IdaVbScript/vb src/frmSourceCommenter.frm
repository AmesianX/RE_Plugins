VERSION 5.00
Begin VB.Form frmSourceCommenter 
   Caption         =   "Source Commenter"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   9360
      TabIndex        =   1
      Top             =   5985
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   5820
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmSourceCommenter.frx":0000
      Top             =   45
      Width           =   10590
   End
End
Attribute VB_Name = "frmSourceCommenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    x = Split(Text1, vbCrLf)
    For i = 0 To UBound(x)
        l = x(i)
        a = InStr(l, " ")
        If a > 0 Then
            b = Mid(l, a)
            c = Mid(l, 1, a - 1)
            x(i) = "/*" & c & "*/" & b
        Else
            x(i) = "//" & x(i)
        End If
    Next
    
    Text1 = Join(x, vbCrLf)
    Text1 = Replace(Text1, ";", "//;")
    
End Sub
