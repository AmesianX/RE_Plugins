VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "FlowChartX TreeLayout sample"
   ClientHeight    =   2655
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   100
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Timer Timer1 
      Left            =   6600
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6960
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuNode 
      Caption         =   "Node menu"
      Visible         =   0   'False
      Begin VB.Menu miAddChild 
         Caption         =   "Add a child node"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuLaunchOriginal 
         Caption         =   "Orginal WinGraph"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public WithEvents udraw As CuDraw
Attribute udraw.VB_VarHelpID = -1
Dim WithEvents idaIpc As CIpc
Attribute idaIpc.VB_VarHelpID = -1

Dim uniqueID As Long
Dim fso As New CFileSystem2
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const last_graph = "c:\ida_last_graph.txt"

'features:
'        loads c:\ida_last_graph.txt or command line file
'        if command line file, save copy as last_graph
'        starts:   C:\Program Files\uDraw(Graph)\bin\uDrawGraph.exe
'        on node selection in uDraw, it can make IDA jump to the function
'        if IDASrvr is installed and running (WM_COPYDATA version)


Function GetValue(e, name, Optional endmark = """")
    On Error Resume Next
    Dim t, x
    t = InStr(e, name)
    
    If t > 0 Then
        t = t + Len(name) + 1
        If endmark = """" Then t = t + 1
        x = InStr(t, e, endmark)
        If x > 0 And x > t Then t = Mid(e, t, x - t)
    End If
    
    GetValue = t
    
End Function



Private Sub Command3_Click()
    Dim n As CNode
    Dim tmp As String
    Dim i As Long
     
    For Each n In udraw.graph.nodes
        If n.va = 0 Then
            i = i + 1
        Else
            tmp = tmp & n.label & ":" & Hex(n.va) & vbCrLf
        End If
    Next
    
    If i > 0 Then
        MsgBox "Could not get VA for " & i & "/" & nodes.Count & " nodes is IDA up and listening?"
    End If
    
    fso.WriteFile "C:\ida_last_adr.txt", tmp
    MsgBox "Save Complete", vbInformation
    
End Sub


Private Sub Form_Load()

    
    Dim dat, entry
    Dim i As Long
    Dim idaUp As Boolean
    Dim fpath As String
    
    On Error Resume Next
    
    Me.Show
    
    Set udraw = New CuDraw
    'udraw.ParentHWND = Picture1.hwnd
    udraw.Initilize Winsock1, Timer1
    
    Set idaIpc = New CIpc
    idaIpc.Listen "Grapher"
    idaIpc.ClientName = "IDA_SERVER"
    idaUp = idaIpc.ClientExists(idaIpc.ClientName)
    
    fpath = Mid(command, InStr(command, """") + 1)
    fpath = Replace(fpath, """", "")
    
    If fso.FileExists(fpath) Then
        dat = fso.ReadFile(fpath)
        If fso.FileExists(last_graph) Then Kill last_graph
        FileCopy fpath, last_graph
    ElseIf fso.FileExists(last_graph) Then
        dat = fso.ReadFile(last_graph)
    Else
        MsgBox "Loads the ida dot graph file passed on on command line from ida or c:\ida_last_graph.txt. Neither was found."
        End
    End If
    
    'FIXME: this works for IDA generated graphs, but not bindiff graphs which are multiline..
    dat = Split(dat, vbCrLf)
        
    Dim node As CNode, nodeB As CNode
    Dim source, target
    
    idaIpc.TimedOut = False
    
    For Each entry In dat
        If Left(entry, 5) = "node:" Then
            
            Set node = udraw.graph.add_node("")
            
            'node: { title: "0" label: "sub_1001AA79" color: 76 textcolor: 73 borderwidth: 10 bordercolor: 82  }
            node.label = GetValue(entry, "label:")
            node.color = GetValue(entry, "color:", " ")
            node.title = GetValue(entry, "title:")
            
            'lets load these on demand to speed up initilization time...actually this wasnt even used!
            'If idaIpc.TimedOut = False And idaIpc.ClientExists(idaIpc.ClientName) Then
            '    node.va = CLng(idaIpc.SendAndRecv("name_va:" & node.label & ":Grapher"))
            'End If
            
        ElseIf Left(entry, 5) = "edge:" Then
            source = GetValue(entry, "sourcename:")
            target = GetValue(entry, "targetname:")
            If udraw.graph.FindNode(source, node) Then
                If udraw.graph.FindNode(target, nodeB) Then
                    udraw.graph.add_edge node, nodeB
                End If
            End If
        End If
    Next
    
    'udraw.SendCommand "app_menu(create_menus([menu_entry(""2"",""menu_label"")]))"
    'udraw.SendCommand "app_menu(activate_menus([""2""]))"
    
    udraw.AddNodeMenu "select children"
    udraw.graph.GenerateGraph
    
End Sub

 



Private Sub Form_Resize()
    On Error Resume Next
    List1.Width = Me.Width - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    udraw.ShutDown
    For i = 0 To 20
        Sleep 50
        DoEvents
    Next
End Sub

Private Sub HScroll1_Change()
    udraw.scaleit HScroll1.Value
    'udraw.layout_improve_all
End Sub

Private Sub idaIpc_SendTimedOut()
    List1.AddItem "IDAIPC Data Send Timed out"
End Sub

Private Sub mnuLaunchOriginal_Click()
    Dim fpath As String
    fpath = App.Path & "\_wingraph32.exe"
    If Not fso.FileExists(fpath) Then
        MsgBox "Original not found must be in cur dir as _wingraph32.exe"
        Exit Sub
    End If
    On Error Resume Next
    Shell fpath & " " & last_graph
End Sub

Private Sub udraw_Message(msg As Variant)
    If msg = "ok" & vbCrLf Then Exit Sub
    List1.AddItem "Message: " & msg
End Sub

Private Sub udraw_NodeSelected(n As CNode)
    List1.AddItem "NodeSelected: " & n.label
    
    idaIpc.Send "jmp_name:" & Trim(n.label)
End Sub

Private Sub udraw_PopupMenuSel(n As CNode, sMenu As String)
    List1.AddItem "Popupmenu: " & sMenu & " on node: " & n.label
    
    Select Case LCase(sMenu)
        Case "select children": n.SelectChildren
    End Select
    
End Sub
