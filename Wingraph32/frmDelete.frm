VERSION 5.00
Begin VB.Form frmDelete 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Nodes"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2325
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   2325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShowAll 
      Caption         =   "Show All"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      TabIndex        =   7
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hide Nodes Below"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hide Nodes Above"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hide Nodes Between"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hide Selected"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "End"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Start"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3000
      Width           =   495
   End
End
Attribute VB_Name = "frmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Text1_HadFocus As Boolean
Public Text2_HadFocus As Boolean

Public StartNode As afNode
Public EndNode As afNode

Private Sub cmdShowAll_Click()
    
    Dim n As afNode
    Dim l As afLink
    
    For Each n In Navig.af.nodes
        For Each l In n.Links
            l.Hidden = False
        Next
        n.Hidden = False
    Next
    
End Sub

Private Sub Command4_Click()
    'hide nodes below
    
    If StartNode Is Nothing Then
        MsgBox "Put cursor in start textbox and click on selected start node", vbInformation
        Exit Sub
    End If
    
    Dim c As New Collection
    Dim n As afNode
    Dim l As afLink
    
    RecursiveLinksUnder StartNode, c
    StartNode.Selected = False
    
    For Each n In c
        If Not n Is StartNode Then
            For Each l In n.Links
                l.Hidden = True
            Next
            n.Hidden = True
        End If
    Next
        
End Sub

Sub RecursiveLinksUnder(n As afNode, c As Collection)
    
    On Error Resume Next
    Dim l As afLink
    Dim nn As afNode
    
    'find all nodes linked to under n
    'how to tell
    
    For Each l In n.Links
        Set nn = l.Dst
        c.Add nn, "oid:" & ObjPtr(nn)
        If Err.Number = 0 And nn.Links.Count > 0 Then
            RecursiveLinksUnder nn, c
        End If
        Err.Clear
    Next
    
End Sub


Private Sub Form_Load()
    FormPos Me, True
    Text1_GotFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormPos Me, True, False
End Sub

Private Sub Text1_GotFocus()
    Text1_HadFocus = True
    Text2_HasFocus = False
End Sub


Private Sub Text2_GotFocus()
    Text2_HadFocus = True
    Text1_HadFocus = False
End Sub

