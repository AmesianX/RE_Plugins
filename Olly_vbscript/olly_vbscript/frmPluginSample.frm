VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmOllyScript 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Olly Debug VB Script Plugin Sample"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   3060
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   1995
      Left            =   0
      TabIndex        =   1
      Top             =   3600
      Width           =   6915
      Begin VB.CommandButton cmdHelp 
         Caption         =   "?"
         Height          =   315
         Left            =   3000
         TabIndex        =   7
         Top             =   1620
         Width           =   375
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load Saved"
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   1620
         Width           =   1395
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   1620
         Width           =   975
      End
      Begin VB.TextBox txtLog 
         Height          =   1395
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   180
         Width           =   6675
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Run Script"
         Height          =   315
         Left            =   5580
         TabIndex        =   2
         Top             =   1620
         Width           =   1215
      End
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
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   6915
   End
End
Attribute VB_Name = "frmOllyScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'dzzie@yahoo.com
'http://sandsprite.com


Public olly As New COlly
Public fso As New clsFileSystem
Private intel As New CIntelliLight


Private IgnoredStartAddress As Long

Private Sub cmdContinue_Click()
    Me.Visible = False
End Sub

Private Sub cmdHelp_Click()
    
    Const msg = "This editor supports a very basic intellisense feature." & vbCrLf & _
                "" & vbCrLf & _
                "After typing a partial token, press ` " & vbCrLf & _
                "" & vbCrLf & _
                "This will can autocomplete a unique token, allow you choose from" & vbCrLf & _
                "multiple matchs, or display all known tokens to choose from." & vbCrLf & _
                "" & vbCrLf & _
                "Arrow up and down to choose, return to select. Function prototype will" & vbCrLf & _
                "be displayed in caption bar upon selection." & vbCrLf & _
                "" & vbCrLf
    
    MsgBox msg, vbInformation
    
End Sub

Private Sub cmdLoad_Click()
    frmScripts.Show 1
End Sub

Private Sub cmdSave_Click()
    Dim n As String
    
    On Error Resume Next
    
    If Len(txtScript) = 0 Then
        MsgBox "Cannot Save Empty Script", vbInformation
        Exit Sub
    End If
    
    n = InputBox("Enter name for script to be saved as")
    
    If Len(n) = 0 Then Exit Sub
    
    Dim sql As String
    
    sql = "Insert into tblScripts (scriptName,scriptText) " & _
          "values('" & Replace(n, "'", "\'") & _
                 "','" & Replace(txtScript, "'", "\'") & "')"
    
    cn.Execute sql
    
    
End Sub

Private Sub Command1_Click()
    On Error GoTo hell
    IgnoredStartAddress = olly.EIP
    sc.AddCode CStr(txtScript.Text)
    Exit Sub
hell:  MsgBox Err.Description
End Sub

Sub HandleBreakpoint()
    On Error GoTo hell
    Dim handler As String
    Dim eip_handler As String
    
    
    If olly.EIP = IgnoredStartAddress Then 'weird
        IgnoredStartAddress = 0
        Exit Sub
    End If
    
    RedrawCpu
    
    handler = "BpxHandler_" & Bpx_Handler
    eip_handler = "BpxHandler_" & Hex(olly.EIP)
    Bpx_Handler = Bpx_Handler + 1
    
    If ProcExists(eip_handler) Then
         sc.Run eip_handler
    ElseIf ProcExists(handler) Then
         sc.Run handler
    ElseIf ProcExists("Default_BPX_Handler") Then
        sc.Run "Default_BPX_Handler"
    ElseIf BpxHandler_Warning Then
        MsgBox "Could not locate a defined sub call for breakpoint action" & vbCrLf & vbCrLf & _
               "Search order for this event would be: " & vbCrLf & vbCrLf & _
                "sub " & eip_handler & vbCrLf & _
                "sub " & handler & vbCrLf & _
                "sub Default_BPX_Handler" & vbCrLf & vbCrLf & _
                "If you did not want to handle this bpx call disableBpxhook", vbInformation
    End If
    
Exit Sub
hell:
        Me.Visible = True
        
        MsgBox "Error in Handle Breakpoint: " & vbCrLf & vbCrLf & _
               "Sc.error: " & sc.Error.Description & vbCrLf & _
               "Sc.line: " & sc.Error.Line, vbInformation
        
        
               
End Sub

Private Function ProcExists(pName As String) As Boolean
    
    On Error GoTo out
    Dim i As Integer
    
    For i = 1 To sc.Procedures.Count
        If LCase(sc.Procedures(i).name) = LCase(pName) Then
            ProcExists = True
            Exit Function
        End If
    Next
    
    Exit Function
out:
End Function


Private Sub Form_Load()

    sc.AddObject "oly", olly, True
    sc.AddObject "fso", fso, True
    sc.AddObject "cn", cn, True
    
    Dim mdb As String
    
    mdb = App.path & "\olly_vbscript.mdb"
    
    If fso.FileExists(mdb) Then
       cn.ConnectionString = Replace("Provider=MSDASQL;Driver={Microsoft Access Driver (*.mdb)};DBQ=____;", "____", mdb)
       cn.Open
    Else
        cmdSave.enabled = False
        cmdLoad.enabled = False
    End If

    Set intel.txt = txtScript
    Set intel.lst = List1
    intel.Load
 

End Sub


 
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Me.Visible = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtScript.Width = Me.Width - 200
    
    Frame1.Top = Me.Height - Frame1.Height - 350
    txtScript.Height = Frame1.Top - 100
    
    Frame1.Width = txtScript.Width
    txtLog.Width = Frame1.Width - 100
    Command1.Left = Frame1.Width - Command1.Width
    
End Sub

 
