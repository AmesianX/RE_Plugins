VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB6 IDASrvr Example"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   10620
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Active IDA Windows"
      Height          =   315
      Left            =   8100
      TabIndex        =   2
      Top             =   2340
      Width           =   2355
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   60
      TabIndex        =   1
      Top             =   2820
      Width           =   10455
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   10515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://support.microsoft.com/kb/176058
'this uses inline subclassing code, I would recommend using a library such as
'my spSubclass or VBaccelerator's subclass lib for stability when running in the IDE.
'ps dont hit end from within IDE or it will crash as subclass isnt cleaned up.

Dim ida As New CIDA

Private Sub Command1_Click()
    
    FindActiveIDAWindows
    
    'If Servers.Count = 0 Then
    '    MsgBox "No server windows have registered themselves.."
    '    Exit Sub
    'End If
    
    ida.ShowServers
    
End Sub

Private Sub Form_Load()
    Dim va As Long
    
    Me.Visible = True
    
    Hook Me.hwnd
    List1.AddItem "Listening for messages on hwnd: " & Me.hwnd

    'this is the original method which connects to the last IDA window opened.
    'support is being added for multiple windows click on the find active windows button.
    If Not ida.FindClient() Then
        List1.AddItem "Could not find IDA Server hwnd."
        Exit Sub
    End If
        
    List1.AddItem "Loaded idb: " & ida.LoadedFile()
    List1.AddItem "NumFuncs: " & ida.NumFuncs()
    
    va = ida.FunctionStart(0)
    List1.AddItem "Func[0].start: " & Hex(va)
    List1.AddItem "Func[0].end: " & Hex(ida.FunctionEnd(0))
    List1.AddItem "Func[0].name: " & ida.FunctionName(0)
    List1.AddItem "1st inst: " & ida.GetAsm(va)
    
    List1.AddItem "Jumping to 1st inst"
    ida.Jump va
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unhook
End Sub
 

