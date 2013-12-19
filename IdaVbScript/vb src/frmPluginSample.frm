VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPluginSample 
   Caption         =   "Misc IDA Functionality"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAsmOnly 
      Caption         =   "Asm Only"
      Height          =   240
      Left            =   8235
      TabIndex        =   33
      Top             =   4635
      Width           =   1680
   End
   Begin VB.CommandButton cmdCopyList 
      Caption         =   "Copy Function List"
      Height          =   255
      Left            =   8280
      TabIndex        =   22
      Top             =   2700
      Width           =   1815
   End
   Begin VB.TextBox txtSelLeng 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   960
      TabIndex        =   18
      Top             =   1500
      Width           =   1035
   End
   Begin VB.TextBox txtSelStart 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   960
      TabIndex        =   16
      Top             =   1020
      Width           =   975
   End
   Begin VB.TextBox txtSelEnd 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   960
      TabIndex        =   17
      Top             =   1260
      Width           =   975
   End
   Begin VB.CheckBox chkListenSocket 
      Caption         =   "Open Cmd Socket"
      Height          =   315
      Left            =   60
      TabIndex        =   13
      Top             =   360
      Width           =   1875
   End
   Begin VB.Frame Frame2 
      Caption         =   " Extra Tools "
      Height          =   5295
      Left            =   0
      TabIndex        =   9
      Top             =   3720
      Width           =   1935
      Begin VB.CommandButton Command3 
         Caption         =   "Copy Xrefs to"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   4080
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Source Comment"
         Height          =   375
         Left            =   135
         TabIndex        =   34
         Top             =   3690
         Width           =   1680
      End
      Begin VB.CommandButton cmdExtractComments 
         Caption         =   "Extract Comments"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   3330
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Search && Comment"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   2925
         Width           =   1695
      End
      Begin VB.CommandButton cmdVBDeclares 
         Caption         =   "VB API Declares"
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   2580
         Width           =   1695
      End
      Begin VB.CommandButton cmdMore 
         Caption         =   "More >>"
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   4860
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Address List"
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   2220
         Width           =   1695
      End
      Begin VB.CommandButton cmdManualPatch 
         Caption         =   "Manual Patch"
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   1860
         Width           =   1695
      End
      Begin VB.CommandButton cmdMemDump 
         Caption         =   "Memory Dump"
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   1140
         Width           =   1695
      End
      Begin VB.CommandButton cmdOffsets 
         Caption         =   "Offset Calc"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdPatch 
         Caption         =   "Import Patch"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   540
         Width           =   1690
      End
      Begin VB.CommandButton cmdOllyImportAPI 
         Caption         =   "Import Olly Calls"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1690
      End
      Begin VB.CommandButton cmdIDAScript 
         Caption         =   "IDA Script"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1500
         Width           =   1690
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Selection Block Tools "
      Height          =   1875
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
      Begin VB.CommandButton cmdOriginals 
         Caption         =   "Show Orig Bytes"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1500
         Visible         =   0   'False
         Width           =   1690
      End
      Begin VB.CommandButton cmdDumpSel 
         Caption         =   "DumpSelection"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Visible         =   0   'False
         Width           =   1690
      End
      Begin VB.CommandButton cmdStronStack 
         Caption         =   "Strs on Stack"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1690
      End
      Begin VB.CommandButton cmdRemPatch 
         Caption         =   "Remove Patch"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   540
         Width           =   1690
      End
      Begin VB.CommandButton cmdUndefine 
         Caption         =   "Undefine Block"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1690
      End
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   4980
      Width           =   7935
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
      Height          =   1575
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3000
      Width           =   7935
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2655
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "n"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Start EA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "End EA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Length"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Func Name"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "MD5"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Selection::"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   24
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "End:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   23
      Top             =   1260
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Length"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   180
      TabIndex        =   15
      Top             =   1500
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Start:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   14
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Asm"
      Height          =   195
      Index           =   1
      Left            =   2100
      TabIndex        =   4
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Function Bytes"
      Height          =   195
      Index           =   0
      Left            =   2100
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmPluginSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'dzzie@yahoo
'http://sandsprite.com

'these cant change while form is modal

'Modeless Forms in VB ActiveX DLL's Don't Display in VC++ Clients
'http://support.microsoft.com/default.aspx?scid=kb;en-us;247791

Private selleng As Long
Private selstart As Long
Private selend As Long

Public pCplugin As CPlugin

Private isLoaded As Boolean

Public Functions As New Collection

 


Private Sub chkListenSocket_Click()
    If chkListenSocket.value = 1 Then
        If pCplugin.IDASCK.isUp Then Exit Sub
        pCplugin.IDASCK.Listen
    Else
        pCplugin.IDASCK.shutdown
    End If
End Sub

Private Sub cmdCopyList_Click()
    Dim tmp() As String
    Dim a As String
    Dim li As ListItem
    Dim i, j
    
    For j = 1 To lv.ColumnHeaders.count
        a = a & lv.ColumnHeaders(j).Text & vbTab
    Next
    
    push tmp, a
    
    For Each li In lv.ListItems
        a = li.Text & vbTab
        For j = 1 To lv.ColumnHeaders.count - 1
            a = a & li.SubItems(j) & vbTab
        Next
        push tmp, a
    Next
        
    Clipboard.Clear
    Clipboard.SetText Join(tmp, vbCrLf)
    
End Sub


Private Sub cmdExtractComments_Click()
    ShowWindow frmExtractComments.hWnd, 1
End Sub

Private Sub cmdManualPatch_Click()
    'frmManualPatch.Show 1, Me
    ShowWindow frmManualPatch.hWnd, 1
End Sub

Private Sub cmdMemDump_Click()

    If ProcessState = 0 Then
        MsgBox "No process Currently Active"
        Exit Sub
    End If
    
    'frmMemDump.Show 1
    ShowWindow frmMemDump.hWnd, 1
    
End Sub

Private Sub cmdMore_Click()
    Dim li As ListItem
    Dim cnt As Long, i As Long
    Dim startPos As Long, endPos As Long
        
    If Me.Width <> 10245 Then
        Me.Width = 10245
        cmdMore.Caption = " << Less"
        cmdDumpSel.Visible = True
        cmdOriginals.Visible = True
        
        If lv.ListItems.count = 0 Then

            cnt = NumFuncs()
                     
            For i = 0 To cnt - 1 'NumFuncs ary 0 based
                Set li = lv.ListItems.Add(, , i)
                startPos = FunctionStart(i)
                endPos = FunctionEnd(i)
                li.SubItems(1) = Hex(startPos)
                li.SubItems(2) = Hex(endPos)
                li.SubItems(3) = endPos - startPos
                li.SubItems(4) = GetFName(startPos)
            Next
            
        End If
    Else
        Me.Width = 2160
        cmdMore.Caption = "More >>"
        cmdDumpSel.Visible = False
        cmdOriginals.Visible = False
    End If
    
    
End Sub

Private Sub cmdOffsets_Click()
    'frmOffsets.Show 1
    ShowWindow frmOffsets.hWnd, 1
End Sub

Private Sub cmdOriginals_Click()
    Dim i As Long
    On Error Resume Next
    
    If selleng > &H1000 Then
        If MsgBox("Selection size is" & selleng & " Are you sure you want to continue?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
       
    Dim ret
    
    For i = selstart To selend - 1
        ret = ret & GetHex(OriginalByte(i)) & " "
    Next
    
    Text2 = ret
    
End Sub

Private Sub cmdPatch_Click()
    frmByteImport.Show 1
End Sub

Private Sub cmdRemPatch_Click()
    Dim i As Long, b As Byte
    On Error Resume Next
   
    If selleng > &H1000 Then
        If MsgBox("Selection size is" & selleng & " Are you sure you want to continue?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If

    For i = selstart To selend
        b = OriginalByte(i)
        PatchByte i, b
    Next
    
    aRefresh
    
End Sub

Private Sub cmdUndefine_Click()
    On Error Resume Next
    Dim i As Long
    
    If selleng > &H1000 Then
        If MsgBox("Selection size is" & selleng & " Are you sure you want to continue?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    For i = selstart To selend
        Undefine i
    Next
      
    aRefresh
    
End Sub

Private Sub cmdDumpSel_Click()

    If selleng > &H1000 Then
        If MsgBox("Selection size is" & selleng & " Are you sure you want to continue?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    DumpData selstart, selleng
    
End Sub

Private Sub cmdVBDeclares_Click()
    frmVBDeclares.Initialize
    ShowWindow frmVBDeclares.hWnd, 1
End Sub

Private Sub Command1_Click()
    ShowWindow frmAdressList.hWnd, 1
End Sub


Private Sub Command2_Click()
    frmTextSearch.Show 1, Me
End Sub

Private Sub Command3_Click()
    
    'seg000:814EB798 wcsncpy
    'seg000:814EE977
    Dim c As Collection
    Dim l As Long
    Dim x
    Dim tmp
    
    
    l = CLng(InputBox("Enter hex addr: ", "", "&h" & Module1.ScreenEA))
    'MsgBox "Using " & Hex(l)
    
    Set c = Module1.GetXRefTo(l)
    'tmp = "To: "
    For Each x In c
        tmp = tmp & Hex(x) & ","
    Next
    
    Clipboard.Clear
    Clipboard.SetText tmp
    
    MsgBox c.count & "References to this address found and copied to clipboard."
    
    tmp = "From: "
   
    Set c = Module1.GetXRefFrom(l)
    For Each x In c
        tmp = tmp & Hex(x) & ","
    Next
    
    MsgBox c.count & " " & tmp
    
    'Module1.AddCodeXRef &H814EE977, &H814EB798
    
    
End Sub

Private Sub Command4_Click()
    frmSourceCommenter.Show 1, Me
End Sub

'Private Sub cmdExport_Click()
'    'hash functions
'    Dim li As ListItem
'
'    On Error Resume Next
'
'    'For Each li In lv.ListItems
'    '    li.SubItems(5) = FunctionHash(CLng(li.Text))
'    'Next
'
'    Dim leng As Long, start As Long
'     Dim buf() As Byte
'    Dim startPos As Long, endPos As Long
'
'
'
'    Dim ado As Object 'New clsAdoKit
'    Dim pth As String
'
'    Set ado = CreateObject("ADOKIT.clsAdoKit")
'
'    If ado Is Nothing Then
'        MsgBox "You do not have Ado kit installed from sandsprite.com"
'        Exit Sub
'    End If
'
'    pth = InputBox("Enter path to database", , "c:\ida.mdb")
'
'    If Not FileExists(pth) Then
'        MsgBox "File not found"
'        Exit Sub
'    End If
'
'    ado.BuildConnectionString 0, pth
'
'    Dim tbl As String
'    tbl = InputBox("Enter table name", , "a")
'
'    For Each li In lv.ListItems
'        leng = li.SubItems(3)
'        start = CLng("&h" & li.SubItems(1))
'        DumpData start, leng
'        ado.Insert tbl, "idb,bytes,disasm,index,leng,fname", tbl, Text1, Text2, li.Text, leng, li.SubItems(4)
'        'If Err.Number > 0 Then MsgBox li.Text & Err.Description
'        'Err.Clear
'    Next
'
'
'
'    MsgBox "Done"
'
'End Sub

Private Sub Form_Load()
    Dim li As ListItem
    Dim cnt As Long, i As Long
    Dim startPos As Long, endPos As Long
    Dim s, e
    Dim shash As String
    
    On Error Resume Next
    
    Me.Width = 2160
    
    s = GetSetting("IDAVBSAMPLE", "SETTINGS", "LEFT", 0)
    e = GetSetting("IDAVBSAMPLE", "SETTINGS", "TOP", 0)
    
    Me.Move s, e
    
    
    cnt = NumFuncs()
    SelBounds selstart, selend
    selleng = selend - selstart
    
    Label1 = "Funcs : " & cnt & "     Block"
    txtSelStart = Hex(selstart)
    txtSelEnd = Hex(selend)
    txtSelLeng = Hex(selleng)
    
    Dim f As CFunction
    On Error Resume Next
    
    For i = 0 To cnt - 1 'NumFuncs ary 0 based
        Set f = New CFunction
        f.Index = i
        startPos = FunctionStart(i)
        endPos = FunctionEnd(i)
        f.StartEA = startPos
        f.EndEA = endPos
        f.Length = endPos - startPos
        f.Name = GetFName(startPos)
        Functions.Add f, "fx:" & f.Name
    Next
    
    If pCplugin.IDASCK.isUp Then chkListenSocket.value = 1
    
    isLoaded = True
    
End Sub

Function FunctionByName(strName) As CFunction
    On Error Resume Next
    Set FunctionByName = Functions("fx:" & Trim(strName))
End Function

Function FunctionIndexByName(strName) As Long
    On Error Resume Next
    Dim f As CFunction
    Set f = Functions("fx:" & Trim(strName))
    If Err.Number <> 0 Then
        FunctionIndexByName = -1
    Else
        FunctionIndexByName = f.Index
    End If
End Function


Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Dim leng As Long, start As Long
    
    leng = Item.SubItems(3)
    start = CLng("&h" & Item.SubItems(1))
        
    DumpData start, leng
    
End Sub

Private Sub cmdStronStack_Click()
    frmStringsOnStack.Show 1
End Sub

Private Sub cmdOllyImportAPI_Click()
    frmDumpFix.Show 1
End Sub

Sub DumpData(ByVal start As Long, ByVal leng As Long)
    
    Text1 = HexDumpBytes(start, leng)
    Text2 = GetAsmRange(start, leng, chkAsmOnly.value)
    
End Sub

'Function FunctionHash(Index As Long) As String
'    Dim buf() As Byte
'    Dim startPos As Long, endPos As Long
'    Dim leng As Long
'
'    startPos = FunctionStart(Index)
'    endPos = FunctionEnd(Index)
'    leng = endPos - startPos
'
'    ReDim buf(leng - 1)
'    GetBytes startPos, buf(0), leng
'
'    FunctionHash = hash.HashBytes(buf)
'
'End Function


Private Sub cmdIDAScript_Click()
    'frmIDAScript.Show 1
    ShowWindow frmIDAScript.hWnd, 1
End Sub

Public Sub Done()
    On Error Resume Next
    Dim f As Form
    For Each f In Forms
        Unload f
    Next
    Unload Me
End Sub



Private Sub txtSelEnd_Change()
 On Error GoTo hell
    
    If Not isLoaded Then Exit Sub
    
    selend = CLng("&h" & txtSelEnd)
    selleng = selend - selstart
        
    If selleng < 0 Then selleng = 0
    
    txtSelLeng = Hex(selleng)
hell:
End Sub

Private Sub txtSelLeng_Change()
 On Error GoTo hell
    
    If Not isLoaded Then Exit Sub
    
    selleng = CLng("&h" & txtSelLeng)
    selend = selstart + selleng
        
    txtSelEnd = Hex(selend)
    
hell:
End Sub

Private Sub txtSelStart_Change()
    On Error GoTo hell
    
    If Not isLoaded Then Exit Sub
    
    selstart = CLng("&h" & txtSelStart)
    selleng = selend - selstart
    
    If selleng < 1 Then selleng = 0
    
    txtSelLeng = Hex(selleng)
    
hell:
End Sub
