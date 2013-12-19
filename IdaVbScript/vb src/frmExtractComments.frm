VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExtractComments 
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7440
      Top             =   360
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Offset"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Comment"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuLoad 
      Caption         =   "Load"
      Begin VB.Menu mnuLoader 
         Caption         =   "From File"
         Index           =   0
      End
      Begin VB.Menu mnuLoader 
         Caption         =   "From CLipboard"
         Index           =   1
      End
      Begin VB.Menu mnuLoader 
         Caption         =   "From Inputbox"
         Index           =   2
      End
      Begin VB.Menu mnuLoader 
         Caption         =   "Help"
         Index           =   4
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy All"
      End
   End
End
Attribute VB_Name = "frmExtractComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ticks As Long

Private Sub mnuLoader_Click(Index As Integer)

    lv.ListItems.Clear
    On Error Resume Next
    
    If Index = 0 Then
        x = gdlg.OpenDialog(AllFiles, , "Load function list", Me.hWnd)
        If Len(x) = 0 Then Exit Sub
        If Not gfso.FileExists(x) Then Exit Sub
        x = gfso.ReadFile(x)
    ElseIf Index = 1 Then
        x = Clipboard.GetText()
        If Len(x) = 0 Then Exit Sub
    ElseIf Index = 2 Then
        x = InputBox("Enter function name to scan")
        If Len(x) = 0 Then Exit Sub
        x = x & vbCrLf
    Else
        Const msg = "Function list is a list of function NAMES one entry per line.\nGenerate this however you want, I extract it from the Chart Refs to tmp file with a tool"
        MsgBox Replace(msg, "\n", vbCrLf), vbInformation
        Exit Sub
    End If
    
    x = Split(x, vbCrLf)
    pb.value = 0
    pb.Max = UBound(x)
    
    Dim f As CFunction
    Dim c As String
    Dim o As Long
    Dim li As ListItem
    Dim l As Long
    Dim i As Long
    Dim base As String
    
    
    
    Timer1.enabled = True
    
    For Each y In x
    
        Err.Clear
        
    
        Set f = frmPluginSample.FunctionByName(y)
        If f Is Nothing Then GoTo next_one
        If Err.Number <> 0 Then GoTo next_one
            
        If InStr(GetAsmCode(f.StartEA), "jmp") > 0 Then GoTo next_one
        
        next_ea = InstructionLength(f.StartEA) + f.StartEA
        If InStr(GetAsmCode(next_ea), "jmp") > 0 Then GoTo next_one
        
        If f.StartEA < 1 Then GoTo next_one
        If f.EndEA < 1 Then GoTo next_one
        If f.EndEA < f.StartEA Then GoTo next_one
        If f.EndEA = f.StartEA Then GoTo next_one
        If f.Length > &H1500 Then GoTo next_one
        If f.Length < 1 Then GoTo next_one
        
        Me.Caption = f.Name & "  " & Hex(f.StartEA) & "-" & Hex(f.EndEA)
        base = Me.Caption
        DoEvents
        Me.Refresh
        DoEvents
        DoEvents
        DoEvents
        
        i = f.StartEA
        ticks = 0
        If Err.Number <> 0 Then GoTo next_one
        
        Do While i < f.EndEA
            
            Me.Caption = base & "   " & Hex(i) & " / " & Hex(f.EndEA - i) & "   " & Now
            Me.Refresh
            DoEvents
            
            Err.Clear
            c = Trim(Get_Comment(i))
            If Len(c) = 0 Then c = Trim(Get_RComment(i))
            
            If Len(c) = 0 Then
                c = Trim(GetAsmCode(i))
                If InStr(c, ";") > 0 Then
                    c = Mid(c, InStr(c, ";") + 1)
                    c = Replace(c, """", Empty)
                    c = Trim(c)
                ElseIf c Like "*call*ds:*" Then
                    c = Trim(c)
                Else
                    c = Empty
                End If
            End If
            
            If Len(c) > 0 Then
                Set li = lv.ListItems.Add(, , Hex(i))
                li.SubItems(1) = c
            End If
            
            l = InstructionLength(CLng(i))
            If l < 1 Then l = 1
            i = i + l
            
            If i > f.StartEA + f.Length + 10 Then Exit Do
            If Err.Number > 0 Then GoTo next_one
            If ticks > 10 Then GoTo next_one
        
            DoEvents
            DoEvents
            Me.Refresh
         Loop
        
next_one:
        DoEvents
        pb.value = pb.value + 1
    Next
        
    Timer1.enabled = False
    MsgBox "Complete " & lv.ListItems.count & " comments found!", vbInformation
    pb.value = 0
    
End Sub

Private Sub Command1_Click()
    MsgBox "Function list is the function name list one per line of all functions to search"
End Sub

Private Sub Command2_Click()
    Dim x
    Dim ea As Long
    On Error GoTo hell
    
    x = InputBox("EA of comment: ", , "&h3201E0FC")
    ea = CLng(x)
    x = Get_RComment(CLng(x))
    MsgBox "RCOmment: " & x
    
    x = GetAsmCode(ea)
    MsgBox x
    
hell:
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    lv.ColumnHeaders(2).Width = lv.Width - lv.ColumnHeaders(2).left - 50
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Jump CLng("&h" & Item.Text)
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuCopyAll_Click()
    Dim li As ListItem
    Dim tmp As String
    
    For Each li In lv.ListItems
        tmp = tmp & li.Text & vbTab & li.SubItems(1) & vbCrLf
    Next
    
    Clipboard.Clear
    Clipboard.SetText tmp
    MsgBox "Copy complete", vbInformation
    
        
        
        
End Sub


Private Sub Timer1_Timer()
    ticks = ticks + 1
End Sub
