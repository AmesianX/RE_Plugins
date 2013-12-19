VERSION 5.00
Begin VB.Form frmStringsOnStack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "stringsOnStack"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   LinkTopic       =   "stringsOnStack"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Safe"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdExample 
      Caption         =   "Ex"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   3360
      Width           =   435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process Text"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   3360
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5475
   End
End
Attribute VB_Name = "frmStringsOnStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExample_Click()
    Text1 = "push 61616161h" & vbCrLf & "mov [eax], 41414141h"
End Sub

Private Sub command1_Click()
    
   On Error Resume Next
    
    
    X = Split(Text1, vbCrLf)
    
    For i = 0 To UBound(X)
       a = Trim(X(i))
       Y = InStrRev(a, ",")
       If Y < 1 Then Y = InStrRev(a, " ")
       If Y < 1 Then Y = 1
       a = Mid(a, Y + 1)
       a = Trim(Replace(a, "h", ""))
       
       For j = 1 To Len(a) Step 2
            z = Mid(a, j, 2)
            Addit c, z
       Next
       
    Next
    
    
    Text1 = c
    
End Sub
 
Function Addit(base, code)
    If code = 0 Then
        base = base & vbCrLf
    ElseIf Check1.Value = 1 Then
        If code > 20 And code < 120 Then
            base = base & Chr("&h" & code)
        Else
            base = base & "."
        End If
    Else
        base = base & Chr("&h" & code)
    End If
End Function
 
  
