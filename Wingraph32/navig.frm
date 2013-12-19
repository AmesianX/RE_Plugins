VERSION 5.00
Object = "{DE173711-6CFE-432E-A95E-F4EF3EE03231}#5.4#0"; "AddFlow5.ocx"
Object = "{69CCB013-9626-43EA-BE8E-4316B96BF952}#2.0#0"; "HFlow2.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EFA5E1B0-F8F3-4374-A076-4EACFA0ABF86}#2.0#0"; "SFlow2.ocx"
Object = "{AFD16587-6577-47B0-A0E7-5EF0E5BDF262}#2.0#0"; "TFlow2.ocx"
Object = "{FC514089-CC6A-4EC1-8D09-07DFA1CF38E2}#1.0#0"; "OFlow.ocx"
Begin VB.Form Navig 
   Caption         =   "AddFlow Navigation tests"
   ClientHeight    =   7755
   ClientLeft      =   1095
   ClientTop       =   945
   ClientWidth     =   12600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7755
   ScaleWidth      =   12600
   Begin VB.CommandButton cmdFindHidden 
      Caption         =   "Find hidden calls"
      Height          =   285
      Left            =   4275
      TabIndex        =   7
      Top             =   180
      Width           =   1410
   End
   Begin VB.Frame fraDebug 
      Caption         =   "IPC Debug Log"
      Height          =   4095
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   9735
      Begin VB.ListBox List1 
         Height          =   3375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   9495
      End
      Begin VB.Label lblCloseDebug 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9270
         TabIndex        =   5
         Top             =   90
         Width           =   375
      End
   End
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   120
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   500
      Left            =   2040
      Top             =   120
   End
   Begin MSComctlLib.ListView lv 
      Height          =   7095
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   12515
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      ToolTipText     =   "Zoom"
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393216
      Min             =   1
      Max             =   1000
      SelStart        =   100
      Value           =   100
   End
   Begin AddFlow5Lib.AddFlow af 
      Height          =   7095
      Left            =   2760
      TabIndex        =   0
      Top             =   600
      Width           =   9735
      _Version        =   327684
      _ExtentX        =   17171
      _ExtentY        =   12515
      _StockProps     =   229
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      ScrollBars      =   3
      Shape           =   1
      LinkStyle       =   0
      Alignment       =   7
      AutoSize        =   3
      ArrowDst        =   3
      ArrowOrg        =   0
      DrawStyle       =   0
      DrawWidth       =   1
      ReadOnly        =   0   'False
      MultiSel        =   -1  'True
      CanDrawNode     =   0   'False
      CanDrawLink     =   0   'False
      CanMoveNode     =   -1  'True
      CanSizeNode     =   0   'False
      CanStretchLink  =   -1  'True
      CanMultiLink    =   -1  'True
      Transparent     =   0   'False
      ShowGrid        =   0   'False
      Hidden          =   0   'False
      Rigid           =   0   'False
      DisplayHandles  =   -1  'True
      AutoScroll      =   -1  'True
      xGrid           =   8
      yGrid           =   8
      xZoom           =   100
      yZoom           =   100
      FillColor       =   16777215
      DrawColor       =   0
      ForeColor       =   0
      BackPicture     =   "navig.frx":0000
      MouseIcon       =   "navig.frx":001C
      AdjustOrg       =   0   'False
      AdjustDst       =   0   'False
      CanReflexLink   =   -1  'True
      SnapToGrid      =   0   'False
      ShowToolTip     =   -1  'True
      ScrollTrack     =   -1  'True
      AllowArrowKeys  =   -1  'True
      ProportionalBars=   -1  'True
      PicturePosition =   9
      LinkCreationMode=   0
      GridStyle       =   0
      ShapeOrientation=   0
      ArrowMid        =   0
      SelectAction    =   0
      GridColor       =   0
      OrthogonalDynamic=   0   'False
      OrientedText    =   0   'False
      EditMode        =   0
      Shadow          =   0
      ShadowColor     =   0
      BackMode        =   1
      Ellipsis        =   0
      SelectionHandleSize=   6
      LinkingHandleSize=   6
      xShadowOffset   =   8
      yShadowOffset   =   8
      CanUndoRedo     =   -1  'True
      UndoSize        =   0
      ShowPropertyPages=   0
      NoPrefix        =   0   'False
      MaxInDegree     =   -1
      MaxOutDegree    =   -1
      MaxDegree       =   -1
      CycleMode       =   0
      LogicalOnly     =   0   'False
      ShowJump        =   0
      SizeArrowDst    =   0
      SizeArrowOrg    =   0
      SizeArrowMid    =   0
      ScrollWheel     =   -1  'True
      RemovePointAngle=   1
      ZeroOriginForExport=   0   'False
      CanFireError    =   0   'False
      RoundedCorner   =   0   'False
      JumpSize        =   0
      RoundedCornerSize=   0
      Autorouting     =   -1  'True
      RouteStartMethod=   0
      RouteGrain      =   8
      RouteMinDistance=   16
      NodeOwnerDraw   =   0   'False
      LinkOwnerDraw   =   0   'False
      OwnerDraw       =   -1  'True
      EditHardReturn  =   0
      Gradient        =   0   'False
      Gradient        =   0
      GradientColor   =   12648447
      Begin OFLOWLib.OFlow OFlow1 
         Left            =   3240
         Top             =   6120
         _Version        =   65536
         _ExtentX        =   1005
         _ExtentY        =   1005
         _StockProps     =   0
         Orientation     =   0
         xMargin         =   600
         yMargin         =   600
         xGrid           =   600
         yGrid           =   600
         NodeSizeRatio   =   50
      End
      Begin TFLOWLib.TFlow TFlow1 
         Left            =   2400
         Top             =   6120
         _Version        =   131072
         _ExtentX        =   1005
         _ExtentY        =   1005
         _StockProps     =   0
         Orientation     =   0
         LayerDistance   =   500
         VertexDistance  =   500
         DrawingStyle    =   0
         xMargin         =   125
         yMargin         =   125
      End
      Begin SYMFLOWLib.SFlow SFlow1 
         Left            =   1440
         Top             =   6120
         _Version        =   131072
         _ExtentX        =   1005
         _ExtentY        =   1005
         _StockProps     =   0
         Distance        =   1000
         SendStepEvent   =   0   'False
         RandomStart     =   -1  'True
         UnmoveableNodesAccepted=   0   'False
         Animation       =   0   'False
         xMargin         =   250
         yMargin         =   250
      End
      Begin MSWinsockLib.Winsock ws 
         Left            =   10920
         Top             =   6120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin HFLOWLib.HFlow HFlow1 
         Left            =   360
         Top             =   6120
         _Version        =   131072
         _ExtentX        =   1005
         _ExtentY        =   1005
         _StockProps     =   0
         Orientation     =   0
         LayerDistance   =   1000
         VertexDistance  =   1000
         LayerWidth      =   0
         xMargin         =   250
         yMargin         =   250
      End
   End
   Begin VB.Label lblNodeCount 
      Caption         =   "Nodes Loaded: 0"
      Height          =   285
      Left            =   45
      TabIndex        =   6
      Top             =   270
      Width           =   2625
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSaveImage 
         Caption         =   "Save Image"
      End
      Begin VB.Menu mnuSaveAddflow 
         Caption         =   "Save As Addflow Graph"
      End
      Begin VB.Menu mnuLoadAddflow 
         Caption         =   "Load Addflow Graph"
      End
      Begin VB.Menu mnuSpacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFile 
         Caption         =   "View Raw uDraw File"
      End
      Begin VB.Menu mnuSpacer4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadLast 
         Caption         =   "Load Last uDraw"
      End
      Begin VB.Menu mnuLoadAnotherUDraw 
         Caption         =   "Load Another uDraw File"
      End
      Begin VB.Menu mnuSpacer5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDebugLog 
         Caption         =   "View Debug Log"
      End
      Begin VB.Menu mnuSpacer6 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMenu 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuZoomIn 
      Caption         =   "Zoom In"
   End
   Begin VB.Menu mnuZoomOut 
      Caption         =   "Zoom Out"
   End
   Begin VB.Menu mnuLaunchOrg 
      Caption         =   "Orginal_Wingraph"
   End
   Begin VB.Menu mnuLayout 
      Caption         =   "Change_Layout"
      Begin VB.Menu mnuHFlow 
         Caption         =   "Heirarchial"
      End
      Begin VB.Menu mnuSFlow 
         Caption         =   "Symetric"
      End
      Begin VB.Menu mnuRadial 
         Caption         =   "Radial"
      End
      Begin VB.Menu mnuTree 
         Caption         =   "Tree"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrthogonal 
         Caption         =   "Orthogonal"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuHideAbove 
         Caption         =   "Hide Nodes Above"
      End
      Begin VB.Menu mnuHideBelow 
         Caption         =   "Hide Nodes Below"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHideSelected 
         Caption         =   "Hide Selected Nodes"
      End
      Begin VB.Menu mnuHideUnSelected 
         Caption         =   "Hide Un-Selected Nodes"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrefixVisible 
         Caption         =   "Prefix Visible Nodes"
      End
      Begin VB.Menu mnuShowAll 
         Caption         =   "Show All Nodes"
      End
      Begin VB.Menu mnuCopyNodeList 
         Caption         =   "Copy Node List"
      End
      Begin VB.Menu mnuSpacer7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRenameNode 
         Caption         =   "Rename Node"
      End
      Begin VB.Menu mnuJumptoNode 
         Caption         =   "Jump To Node in IDA"
      End
   End
End
Attribute VB_Name = "Navig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dlog As String
Private fpath As String
Private zoomIncrement
Dim graph As New CGrapher
Dim done_received As Boolean
Dim time_out As Boolean
Dim dlg As New clsCmnDlg
Public triedtoInit As Boolean

Dim ida As New CIDAScript
Dim WithEvents ipc As CIpc
Attribute ipc.VB_VarHelpID = -1

'Changelog:
'   12-16-12 changed over to use WM_COPYDATA instead of udp sockets


Private Sub ipc_DataReceived(msg As String)
    List1.AddItem "Ipc Data: " & msg
End Sub

Private Sub ipc_DataSend(msg As String, isBlocking As Boolean)
    List1.AddItem "Ipc Send: " & msg & " Blocking: " & isBlocking
End Sub

Private Sub ipc_Error(msg As String)
    List1.AddItem "IPC Error: " & msg
End Sub

Private Sub ipc_SendTimedOut()
    List1.AddItem "Ipc Timeout"
End Sub

'Private Sub StartTimeoutMonitor()
'    time_out = False
'    tmrTimeout.Enabled = True
'End Sub

Private Sub af_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuPopup
    End If
End Sub

Private Sub lblCloseDebug_Click()
    fraDebug.Visible = False
    mnuDebugLog.Checked = False
End Sub

Private Sub lv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuCopyNodeList_Click()
    Dim li As ListItem
    Dim tmp As String
    Dim x
    
    For Each li In lv.ListItems
        x = li.Text
        If InStr(x, ":") > 0 Then
            x = Mid(x, 1, InStr(x, ":"))
            x = Replace(x, ":", Empty)
        End If
        tmp = tmp & x & vbCrLf
    Next
    
    Clipboard.Clear
    Clipboard.SetText tmp
    MsgBox "Copy complete", vbInformation
        
End Sub

Private Sub mnuHFlow_Click()
    HFlow1.layout af
    tmrRefresh.Enabled = True
End Sub

Private Sub mnuJumptoNode_Click()
    If af.SelectedNode Is Nothing Then
        MsgBox "Select a node first. You can also trigger this by double clicking on a node in the graph view", vbInformation
        Exit Sub
    End If
    af_DblClick
End Sub

Private Sub mnuLoadAddflow_Click()
    f = dlg.OpenDialog(AllFiles, , , Me.hwnd)
    If Len(f) > 0 Then
        af.SelectAll
        af.SelNodes.Clear
        af.LoadFile f
    End If
End Sub

Private Sub mnuLoadAnotherUDraw_Click()
    f = dlg.OpenDialog(AllFiles, , , Me.hwnd)
    If Len(f) > 0 Then
        af.SelectAll
        af.SelNodes.Clear
        Form_Load
        startup f
    End If
End Sub

Private Sub mnuOrthogonal_Click()
    OFlow1.layout af
    tmrRefresh.Enabled = True
End Sub

Private Sub mnuRenameNode_Click()
    If af.SelectedNode Is Nothing Then
        MsgBox "Select a node first", vbInformation
        Exit Sub
    End If
    Dim new_name As String
    new_name = InputBox("Enter a new name for node " & Trim(af.SelectedNode.Text), , Trim(af.SelectedNode.Text))
    If Len(Trim(new_name)) = 0 Then Exit Sub
    new_name = Replace(new_name, " ", "_")
    If Not RenameNode(af.SelectedNode, new_name) Then
        MsgBox "Rename Failed See debug log for more info", vbInformation
    End If
End Sub

Private Sub mnuSaveAddflow_Click()
    f = dlg.SaveDialog(AllFiles, , , , Me.hwnd)
    If Len(f) > 0 Then
        af.SaveFile f
    End If
End Sub

Private Sub mnuSaveImage_Click()
    f = dlg.SaveDialog(AllFiles, , , , Me.hwnd)
    If Len(f) > 0 Then
        af.SaveImage afTypeMediumFile, afWMF, f
    End If
End Sub

'Private Sub mnuTree_Click()
'    'TFlow1.VertexDistance = 2000
'    TFlow1.DrawingStyle = Layered
'    TFlow1.layout af
'    tmrRefresh.Enabled = True
'End Sub

Private Sub mnuRadial_Click()
    'TFlow1.VertexDistance = 2000
    TFlow1.DrawingStyle = Radial
    TFlow1.layout af
    tmrRefresh.Enabled = True
End Sub

Private Sub mnuSFlow_Click()
    SFlow1.layout af
    tmrRefresh.Enabled = True
End Sub

Private Sub mnuHideAbove_Click()

    Dim startnode As afNode
    Dim c As New Collection
    Dim n As afNode
    Dim l As afLink
    
    Set startnode = af.SelectedNode
    If startnode Is Nothing Then Exit Sub
    
    RecursiveLinks startnode, c, False
    
    For Each n In c
        If Not n Is startnode Then
            For Each l In n.Links
                l.Hidden = True
            Next
            n.Hidden = True
        End If
    Next
    
    On Error Resume Next
    Dim nn As afNode
    
    Set c = New Collection
    RecursiveLinks startnode, c 'get list of nodes below start these are only ones to be visible
    
    For Each n In af.nodes 'cycle through all nodes
        Set nn = c("oid:" & ObjPtr(n)) 'if this node isnt in the below list, hide it
        If Err.Number <> 0 Then
            For Each l In n.Links
                l.Hidden = True
            Next
            n.Hidden = True
        End If
        Err.Clear
    Next
    
    
End Sub

Private Sub mnuHideSelected_Click()
    Dim n As afNode
    Dim l As afLink
    
    For Each n In Navig.af.nodes
        If n.Selected Then
            For Each l In n.Links
                l.Hidden = True
            Next
            n.Hidden = True
            n.Selected = False
        End If
    Next
    
End Sub

Private Sub mnuHideUnSelected_Click()
    Dim n As afNode
    Dim l As afLink
    
    For Each n In Navig.af.nodes
        If Not n.Selected Then
            For Each l In n.Links
                l.Hidden = True
            Next
            n.Hidden = True
            n.Selected = False
        End If
    Next
    
End Sub



Private Sub mnuPrefixVisible_Click()
    
    Dim n As afNode
    Dim pf As String
    Dim cmd As String
    Dim new_name As String
    Dim li As ListItem
    
    pf = InputBox("Enter prefix to add to all function names")
    pf = Replace(pf, " ", "_")
    
    If Len(pf) = 0 Then Exit Sub
    
    'if the rename isnt successful, then future jumps based on the node text will fail...
    '12.16.11 - ida rename is confirmed with error code now..
    For Each n In af.nodes
        If n.Hidden = False Then
            new_name = pf & Trim(n.Text)
            RenameNode n, new_name
        End If
    Next
    
    
    
    
End Sub


Function RenameNode(n As afNode, new_name As String) As Boolean
    
    On Error GoTo hell
    
        Dim cmd As String
        Dim li As ListItem
        
        If ida.Rename(Trim(n.Text), new_name) Then
            For Each li In lv.ListItems
                If li.Text = n.Text Then
                    li.Text = new_name
                    Exit For
                End If
            Next
            n.Text = new_name
            RenameNode = True
        End If
            
        Exit Function
hell:
        List1.AddItem "Error in RenameNode(" & new_name & ") Desc: " & Err.Description & vbCrLf
        
End Function



'Function RenameNode(n As afNode, new_name As String) As Boolean
'
'    On Error GoTo hell
'
'        Dim cmd As String
'        Dim li As ListItem
'
'        cmd = "newname " & Trim(n.Text) & " " & new_name
'        dlog = dlog & cmd & vbCrLf
'
'        done_received = False
'        StartTimeoutMonitor
'
'        ws.SendData cmd
'
'        Do While Not done_received
'            Sleep 100
'            If time_out Then Exit Do
'            DoEvents
'            DoEvents
'        Loop
'
'        tmrTimeout.Enabled = False
'
'        If Not time_out Then
'            For Each li In lv.ListItems
'                If li.Text = n.Text Then li.Text = new_name
'            Next
'            n.Text = new_name
'            RenameNode = True
'        End If
'
'        Exit Function
'hell:
'        dlog = dlog & "Error in RenameNode(" & new_name & ") Desc: " & Err.Description & vbCrLf
'
'End Function


Private Sub mnuShowAll_Click()
    
    Dim n As afNode
    Dim l As afLink
    
    For Each n In Navig.af.nodes
        For Each l In n.Links
            l.Hidden = False
        Next
        n.Hidden = False
    Next
    
End Sub

Private Sub mnuHideBelow_Click()
    
    Dim startnode As afNode
    Dim c As New Collection
    Dim n As afNode
    Dim l As afLink
    
    Set startnode = af.SelectedNode
    If startnode Is Nothing Then Exit Sub
    
    RecursiveLinks startnode, c
    
    For Each n In c
        If Not n Is startnode Then
            For Each l In n.Links
                l.Hidden = True
            Next
            n.Hidden = True
        End If
    Next
        
End Sub


Sub RecursiveLinks(n As afNode, c As Collection, Optional Under As Boolean = True)
    
    On Error Resume Next
    Dim l As afLink
    Dim nn As afNode
    
    For Each l In n.Links
        If Under Then Set nn = l.Dst Else Set nn = l.Org
        c.Add nn, "oid:" & ObjPtr(nn)
        If Err.Number = 0 And nn.Links.Count > 0 Then
            RecursiveLinks nn, c, Under
        End If
        Err.Clear
    Next
    
End Sub

Private Sub af_DblClick()
        
    On Error Resume Next
    
    If af.SelectedNode Is Nothing Then Exit Sub
    
    ida.JumpName Trim(af.SelectedNode.Text)
    
    'cmd = "jmpfunc " & Trim(af.SelectedNode.Text)
    'dlog = dlog & cmd & vbCrLf
    'ws.SendData cmd
    'If Err.Number <> 0 Then dlog = dlog & Err.Description & vbCrLf

End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set ipc = Nothing
    Set ida = Nothing
    FormPos Me, True, True
    For Each f In Forms
        Unload f
    Next
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Dim n As afNode
    Set n = Item.Tag
    n.EnsureVisible
    af.SelectedNode.Selected = False
    n.Selected = True
    Call af_Click
End Sub


Private Sub mnuLaunchOrg_Click()
    
    On Error Resume Next
    
    'remove that shitty library color you cant change...
    tmp = fso.ReadFile(fpath)
    fso.WriteFile fpath, Replace(tmp, "color: 75 textcolor: 73", "color: white textcolor: black")
    
    If Not isIDE() Then
        Shell App.path & "\wingraph32_.exe """ & fpath & """", vbNormalFocus
    Else
        Shell "c:\program files\IDA\" & "\wingraph32_.exe """ & fpath & """", vbNormalFocus
    End If
    
End Sub

Private Sub af_Click()
    On Error Resume Next
    af.StretchingPoint = af.SelectedNode
End Sub

Private Sub ExitMenu_Click()
    End
End Sub




Private Sub Form_Load()

    Dim a As Long, b As Long
    
    FormPos Me, True
    af.DisplayHandles = True
    zoomIncrement = 150
    
    a = InStr(Command, """")
    If a > 0 Then
        pth = Mid(Command, a + 1)
        pth = Replace(pth, """", "")
    End If
        
    If Len(Command) = 0 Then
        pth = GetSetting("wingraph32", "section", "lastfile", "")
    Else
        SaveSetting "wingraph32", "section", "lastfile", pth
    End If
    
    ida.Initialize Me.hwnd, "Wingraph32"
    List1.AddItem "Listening on hwnd: " & Me.hwnd & " (0x" & Hex(Me.hwnd) & ")"
    
    If ida.isUp Then
        List1.AddItem "IDA Server Up hwnd=" & ida.ipc.ClientHWND & " (0x" & Hex(ida.ipc.ClientHWND) & ")"
        List1.AddItem "IDB: " & ida.LoadedFile
    End If
    
    Set ipc = ida.ipc
    
    lv.ListItems.Clear
    lv.ColumnHeaders(1).Width = lv.Width - 100
    
    fpath = pth
    dlog = "Commandline: " & Command & vbCrLf & "File Path:" & pth & vbCrLf
    startup pth
    lblNodeCount = "Loaded Nodes: " & lv.ListItems.Count
    
End Sub

Sub startup(pth)

    With graph
        Set .afControl = af
        Set .layout = HFlow1
        Set .fxLv = lv
        .LoadFile CStr(pth)
    End With
    
     With af
        '.DisplayHandles = False
        .EditMode = 0
        '.yZoom = 0
        '.xZoom = 0
        .ScrollWheel = True
        '.ScrollBars = afNoScroll
        .Autorouting = True
        .AutoScroll = True
        .AllowArrowKeys = True
        .MultiSel = True
        .ProportionalBars = True
        .SelectMode = True
    End With
    
    'af.SelectAll
    'af.AutoScroll = True
    
    'HFlow1.layout af
 
    'mnuReconnect_Click

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        af.Move af.Left, af.Top, Me.Width - 100 - af.Left, Me.Height - af.Top - 700
        lv.Height = af.Height
    End If
End Sub

Private Sub mnuDebugLog_Click()
    'Form1.showit dlog
    mnuDebugLog.Checked = Not mnuDebugLog.Checked
    fraDebug.Visible = mnuDebugLog.Checked
End Sub





Private Sub mnuLoadLast_Click()
    pth = GetSetting("wingraph32", "section", "lastfile", "")
    On Error Resume Next
    Shell "notepad.exe " & pth, vbNormalFocus
End Sub



'Private Sub mnuReconnect_Click()
'    On Error Resume Next
'    ws.Close
'    ws.Connect "127.0.0.1", 2222
'End Sub



Private Sub mnuViewFile_Click()
    On Error Resume Next
    If Len(fpath) = 0 Then
        MsgBox "File path not set", vbInformation
        Exit Sub
    End If
    If fso.FileExists(fpath) Then
        Shell "notepad.exe " & fpath, vbNormalFocus
    Else
        MsgBox "File not found: " & fpath
    End If
End Sub

Private Sub mnuZoomIn_Click()
    On Error Resume Next
    Slider1.value = Slider1.value + 10
    Slider1_Click
End Sub


Private Sub mnuZoomOut_Click()
    Slider1.value = Slider1.value - 10
    Slider1_Click
End Sub

Private Sub Slider1_Click()
    On Error Resume Next
    x = af.xScroll
    y = af.yScroll
    
    af.xZoom = Slider1.value
    af.yZoom = Slider1.value
    af.SelectedNode.EnsureVisible
    
End Sub


Private Sub tmrRefresh_Timer()
    
    'tmrRefresh.Enabled = False
    DoEvents
    TimerProc
    
End Sub



'Private Sub tmrTimeout_Timer()
'    time_out = True
'    tmrTimeout.Enabled = False
'    dlog = dlog & "tmrTimeout triggered" & vbCrLf
'End Sub
'
'Private Sub ws_Connect()
'    dlog = dlog & "connected" & vbCrLf
'End Sub
'
'Private Sub ws_DataArrival(ByVal bytesTotal As Long)
'    On Error Resume Next
'    Dim x As String
'
'    ws.GetData x
'
'    If InStr(1, x, "ids:done", vbTextCompare) > 0 Then
'        done_received = True
'        x = Replace(x, "ids:done", Empty)
'    End If
'
'    If Len(x) > 0 And Trim(x) <> vbCrLf Then
'        dlog = dlog & "Data Arrived: " & x & vbCrLf
'    End If
'
'End Sub

Function isIDE() As Boolean
    On Error GoTo hell
    Debug.Print 1 / 0
    isIDE = False
    Exit Function
hell:
    isIDE = True
End Function



'Sub HideOrphanNodes() 'if a node has no visible links to it, hide it.
'    Dim n As afNode
'    Dim l As afLink
'    Dim lVisible As Boolean
'
'    For Each n In af.nodes
'        If n.Links.Count > 0 Then
'            lVisible = False
'            For Each l In n.Links
'                If l.Hidden = False Then lVisible = True
'            Next
'            If Not lVisible Then n.Hidden = True
'        End If
'    Next
'
'End Sub


'Sub delParents(nn As afNode)
'
'    Dim n As afNode
'    Debug.Print nn.Text & " Out:" & nn.OutLinks.Count & " In:" & nn.InLinks.Count
'
'    'if n.OutLinks =1 and n.InLinks =1 then
'    For Each l In nn.InLinks
'        Set n = nn.GetLinkedNode(l)
'        delParents n
'        Debug.Print n.Text & " " & n.OutLinks.Count
'        If n.InLinks.Count = 1 Then n.Marked = True
'    Next
 '
 '
'
'End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    'MsgBox KeyCode
'
'    Select Case KeyCode
'        Case 46 'delete key
'               ' If af.SelectedNode Is Nothing Then Exit Sub
'               ' Dim n As afNode
'               ' Dim l As afLink
'               ' delParents af.SelectedNode
'               ' af.SelectedNode.Marked = True
'               '  af.DeleteMarked
'        Case 90: af.Undo
'        Case 89: af.Redo
'    End Select
'End Sub

    
    'If af.SelectedNode Is Nothing Then Exit Sub
   '
   ' If Text1_HadFocus Then
   '     Text1 = af.SelectedNode.Text
   '     Set StartNode = af.SelectedNode
   ' ElseIf Text2_HadFocus Then
   '     Text2 = af.SelectedNode.Text
   '     Set EndNode = af.SelectedNode
   ' End If
