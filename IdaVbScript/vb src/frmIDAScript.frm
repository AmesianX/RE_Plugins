VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmIDAScript 
   Caption         =   "IDA VB Script Interface"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Copy"
      Height          =   255
      Index           =   1
      Left            =   5340
      TabIndex        =   11
      Top             =   3900
      Width           =   495
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   1620
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSScriptControlCtl.ScriptControl sc 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.TextBox txtScript 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   0
      Width           =   7155
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   3540
      Width           =   7275
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   255
         Index           =   0
         Left            =   5340
         TabIndex        =   10
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   315
         Left            =   3120
         TabIndex        =   7
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save New"
         Height          =   315
         Left            =   4140
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load Saved"
         Height          =   315
         Left            =   1980
         TabIndex        =   5
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdExec 
         Caption         =   "Execute"
         Height          =   315
         Left            =   5940
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtLog 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   660
         Width           =   7095
      End
      Begin VB.CommandButton cmdPrototypes 
         Caption         =   "Prototypes"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "?"
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmIDAScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intel As New CIntelliLight
Dim WithEvents extender As CTxtExtender
Attribute extender.VB_VarHelpID = -1

Dim fso As New clsFileSystem
Dim bin As New CBinWrite

Public cn As New Connection

Public selectedID As Long
Public selectedName As String

Private Sub cmdClear_Click(Index As Integer)
    Select Case Index
        Case 0: txtLog = Empty
        Case 1: Clipboard.Clear: Clipboard.SetText txtLog
    End Select
End Sub

Private Sub cmdExec_Click()
    On Error Resume Next
    sc.AddCode txtScript.Text
End Sub

 

Private Sub cmdPrototypes_Click()
    On Error Resume Next
    Shell "notepad """ & App.path & "\protos.txt""", vbNormalFocus
End Sub

Private Sub cmdUpdate_Click()
    
    If selectedID = 0 Then
        MsgBox "Current script id = 0 cannot update"
        Exit Sub
    End If
    
    On Error Resume Next
    
    If Not MsgBox("Are you sure you want to update " & selectedName & " ?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    cn.Execute "Update tblscripts set scripttext='" & Replace(txtScript, "'", "''") & "'" & " where autoid=" & selectedID
    
    If Err.Number > 0 Then MsgBox Err.Description
    
    MsgBox selectedName & " updated successfully"
    
End Sub

Private Sub command1_Click()

    Const Msg = "This editor supports a very basic intellisense feature." & vbCrLf & _
                "" & vbCrLf & _
                "After typing a partial token, press ` " & vbCrLf & _
                "" & vbCrLf & _
                "This will can autocomplete a unique token, allow you choose from" & vbCrLf & _
                "multiple matchs, or display all known tokens to choose from." & vbCrLf & _
                "" & vbCrLf & _
                "Arrow up and down to choose, return to select. Function prototype will" & vbCrLf & _
                "be displayed in caption bar upon selection." & vbCrLf & _
                "" & vbCrLf
    
    MsgBox Msg, vbInformation
    
End Sub

 


Private Sub cmdLoad_Click()
    frmLoadScript.Show 1, Me
End Sub

Private Sub cmdSave_Click()
    Dim X As String
    On Error Resume Next
    X = InputBox("Enter name for current script")
    cn.Execute "Insert into tblScripts (scriptName,ScriptText) values('" & X & "','" & Replace(txtScript.Text, "'", "''") & "')"
    If Err.Number = 0 Then
        MsgBox "Saved Successfully", vbInformation
    Else
        MsgBox "Error: " & Err.Description
    End If
    
    selectedName = X
    selectedID = cn.Execute("Select top 1 autoid from tblScripts order by autoid desc")!autoid
    Me.Caption = "New Script ID=" & selectedID
    
End Sub

 

Private Sub Form_Load()

    sc.AddObject "ida", cIda, True
    sc.AddObject "fso", fso, True
    sc.AddObject "bin", bin, True
    
    Set intel.lst = List1
    Set intel.txt = txtScript
    intel.Load
    
    Set extender = New CTxtExtender
    Set extender.mTextBox = txtScript
    
    extender.AutoIndent = True
    extender.AddAccelerators = True
    
    Dim pth As String
    
    pth = App.path & "\db1.mdb"
    
    If Not Dir(pth) <> "" Then
        MsgBox "Database not found!"
        Exit Sub
    End If
        
   Dim s As String
    
    s = "Provider=MSDASQL;Driver={Microsoft Access Driver (*.mdb)};DBQ=[dbpath];"
    s = Replace(s, "[dbpath]", pth)

    cn.ConnectionString = s
    cn.Open
    
 
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtScript.Width = Me.Width - 50
    Frame1.Width = txtScript.Width
    txtLog.Width = txtScript.Width - 50
    txtScript.Height = Me.Height - 200 - Frame1.Height
    Frame1.Top = Me.Height - Frame1.Height - 100
End Sub
 

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    cn.Close
End Sub

Private Sub sc_Error()
    On Error Resume Next
    MsgBox "Line: " & sc.Error.Line & " Desc:" & sc.Error.Description, vbInformation, "Script Error"
End Sub
 
 
 
 
