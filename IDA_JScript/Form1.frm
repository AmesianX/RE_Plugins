VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Form1 
   Caption         =   "IDA JScript - http://sandsprite.com"
   ClientHeight    =   7020
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Log Window and Output Pane"
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   9975
      Begin VB.CheckBox Check1 
         Caption         =   "Show Debug Log"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Run Script"
         Height          =   465
         Left            =   8400
         TabIndex        =   4
         Top             =   1800
         Width           =   1320
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   1020
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   8865
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   9615
      End
      Begin VB.Label lblIDB 
         Caption         =   "Current IDB (null)"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   1860
         Width           =   6135
      End
   End
   Begin MSScriptControlCtl.ScriptControl sc 
      Left            =   135
      Top             =   4410
      _ExtentX        =   1005
      _ExtentY        =   1005
      Language        =   "JScript"
   End
   Begin Project1.ucScint txtJS 
      Height          =   4245
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   7488
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuOpenScript 
         Caption         =   "Open File"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuSpacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectIDAInstance 
         Caption         =   "Select Active IDA Instance"
      End
      Begin VB.Menu mnuSHellExt 
         Caption         =   "Register .idajs Shell Extension"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ida As New CIDAScript

Private Sub Check1_Click()
    List1.Visible = CBool(Check1.Value)
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    
    Text1 = Empty
    ida.WriteFile App.path & "\lastScript.txt", txtJS.Text
    
    If Not ida.isUp Then
        Text1 = "IDA Server not found"
        lblIDB.Caption = "Current IDB: (null)"
        Exit Sub
    End If
    
    sc.Reset
    sc.AddObject "ida", ida, True
    sc.AddObject "fso", ida, True  'parlor trick to break up intellisense list into smaller segments..
    
    Const wrappers = "function h(x){ return ida.intToHex(x);};" & _
                     "function t(x){ ida.t(x);};" & _
                     "function alert(x){ ida.alert(x);};" & _
                     vbCrLf
    
    sc.AddCode wrappers & txtJS.Text
    
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    
    Dim hwnd As Long
    Dim idb As String
    Dim windows As Long
    
    FormPos Me, True
    Me.Visible = True
    
    txtJS.WordWrap = True
    txtJS.LineIndentGuide = True
    txtJS.Folding = True
    txtJS.AutoCompleteOnCTRLSpace = False
    
    List1.AddItem "Listening on hwnd: " & Me.hwnd & " (0x" & Hex(Me.hwnd) & ")"
    
    If fso.FileExists(Command) Then
        txtJS.Text = ida.ReadFile(Command)
    ElseIf ida.FileExists(App.path & "\lastScript.txt") Then
        txtJS.Text = ida.ReadFile(App.path & "\lastScript.txt")
    End If
    
    windows = ida.ipc.FindActiveIDAWindows()
    If windows = 0 Then
        List1.AddItem "No open IDA Windows detected. Use Tools menu to connect latter."
    ElseIf windows = 1 Then
        ida.ipc.RemoteHWND = ida.ipc.Servers(1)
        idb = ida.LoadedFile
        List1.AddItem "IDA Server Up hwnd=" & ida.ipc.RemoteHWND & " (0x" & Hex(ida.ipc.RemoteHWND) & ")"
        List1.AddItem "IDB: " & idb
        lblIDB = "Current IDB: " & fso.FileNameFromPath(idb)
    Else
        hwnd = Form2.SelectIDAInstance()
        If hwnd <> 0 Then
            ida.ipc.RemoteHWND = hwnd
            idb = ida.LoadedFile
            List1.AddItem "IDA Server Up hwnd=" & ida.ipc.RemoteHWND & " (0x" & Hex(ida.ipc.RemoteHWND) & ")"
            List1.AddItem "IDB: " & idb
            lblIDB = "Current IDB: " & fso.FileNameFromPath(idb)
        End If
    End If
    
    List1.Move Text1.Left, Text1.Top, Text1.Width, Text1.Height
    Text1 = "Built in classes: ida. fso." & vbCrLf & "global functions alert(x), h(x) [int to hex], t(x) [append textbox with x]"
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    txtJS.Width = Me.Width - txtJS.Left - 140
    txtJS.Height = Me.Height - txtJS.Top - Frame1.Height - 450
    Frame1.Width = Me.Width - Frame1.Left - 140
    Frame1.Top = txtJS.Top + txtJS.Height + 50
    Text1.Width = Frame1.Width - Text1.Left - 140
    List1.Move Text1.Left, Text1.Top, Text1.Width, Text1.Height
    List1.Width = Text1.Width
    Command1.Left = Frame1.Width - Command1.Width - 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    FormPos Me, True, True
    If Len(txtJS.Text) > 2 Then ida.WriteFile App.path & "\lastScript.txt", txtJS.Text
End Sub

'Private Sub ipc_DataReceived(msg As String)
'
'End Sub
'
'Private Sub ipc_DataSend(msg As String)
'    List1.AddItem "Ipc Send: " & msg
'End Sub

Private Sub mnuOpenScript_Click()
    
    Dim fpath As String
    fpath = dlg.OpenDialog(AllFiles, , , Me.hwnd)
    If Len(fpath) = 0 Then Exit Sub
    
    txtJS.Text = ida.ReadFile(fpath)
    
End Sub

Private Sub mnuSaveAs_Click()
    
    Dim fpath As String
    Dim ext As String
    ext = ".idajs"
    
    fpath = dlg.SaveDialog(AllFiles)
    If Len(fpath) = 0 Then Exit Sub
    If VBA.Right(fpath, Len(ext)) <> ext Then fpath = fpath & ext
    
    fso.WriteFile fpath, txtJS.Text
    
End Sub

Private Sub mnuSelectIDAInstance_Click()
    Dim hwnd As Long
    Dim idb As String
    
    On Error Resume Next
    hwnd = Form2.SelectIDAInstance()
    If hwnd = 0 Then Exit Sub
    
    ida.ipc.RemoteHWND = hwnd
    idb = ida.LoadedFile()
    lblIDB = "Current IDB: " & fso.FileNameFromPath(idb)
    
End Sub

Private Sub mnuSHellExt_Click()
    
    Dim homedir As String
    
    homedir = App.path & "\IDA_JScript.exe"
    If Not FileExists(homedir) Then Exit Sub
    cmd = "cmd /c ftype IDAJS.Document=""" & homedir & """ %1 && assoc .idajs=IDAJS.Document"
    
    On Error Resume Next
    Shell cmd, vbHide
    
'    Dim wsh As Object 'WshShell
'    Set wsh = CreateObject("WScript.Shell")
'    If Not wsh Is Nothing Then
'        wsh.RegWrite "HKCR\IDAJS.Document\DefaultIcon\", homedir & ",0"
'    End If
    
    
End Sub



Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Private Sub sc_Error()
    '-1 is for the extra line we add silently for wrappers
    Text1 = "Error on line: " & (sc.Error.line - 1) & vbCrLf & sc.Error.Description
    txtJS.GotoLine sc.Error.line
End Sub

Private Sub txtJS_AutoCompleteEvent(className As String)

    If className = "fso" Then
        txtJS.ShowAutoComplete "exec readfile writefile appendfile fileexists deletefile getClipboard setClipboard"
    ElseIf className = "ida" Then
        'do i want to break these up into smaller chunks for intellisense?
        txtJS.ShowAutoComplete "imagebase loadedfile jump patchbyte originalbyte readbyte inttohex refresh() " & _
                               "numfuncs() functionstart functionend functionname getasm instsize xrefsto " & _
                               "xrefsfrom undefine getname jumprva " & _
                               "hideea showea hideblock showblock removename setname makecode " & _
                               "getcomment addcomment addcodexref adddataxref delcodexref deldataxref " & _
                               "funcindexfromva funcvabyname nextea prevea patchstring makestr makeunk " & _
                               "renamefunc"
    End If
        
    'divide up into these classes for intellise sense cleanliness?
    'ui -> jump refresh() hideea showea hideblock showblock getcomment addcomment loadedfile
    'refs -> getrefsto getrefsfrom addcodexref adddataxref delcodexref deldataxref
    'func -> numfuncs() functionstart functionend functionname getname removename setname funcindexfromva funcvabyname
    'code -> imagebase undefine makecode getasm instsize patchbyte orginalbyte readbyte nextea
    
End Sub
