VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{9A143468-B450-48DD-930D-925078198E4D}#1.0#0"; "hexed.ocx"
Begin VB.Form Form1 
   Caption         =   "PECarve http://sandsprite.com"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lv2 
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      Top             =   7200
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
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
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Height          =   375
      Left            =   12480
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin rhexed.HexEd he 
      Height          =   6615
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   11668
   End
   Begin MSComctlLib.ListView lv 
      Height          =   6615
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   11668
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Offset"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   10800
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LoadedFile As String

Private Sub cmdExtract_Click()
    Dim li As ListItem
    Dim pf As String
    Dim start As Long
    Dim leng As Long
    Dim b() As Byte
    Dim f As Long, f1 As Long
    Dim fp As String
    Dim li2 As ListItem
    
    'On Error Resume Next
    lv2.ListItems.Clear
    
    pf = fso.GetParentFolder(LoadedFile) & "\pecarve"
    If Not fso.FolderExists(pf) Then MkDir pf
    
    f = FreeFile
    Open LoadedFile For Binary As f
        
    For i = 1 To lv.ListItems.Count
        Set li = lv.ListItems(i)
        Set li2 = GetNext(i + 1)
        
        start = CLng("&h" & li.Text)
        
        If Not li2 Is Nothing Then
            leng = CLng("&h" & li2.Text) - start - 1
        Else
            leng = LOF(f) - start
        End If
        
        lv2.ListItems.Add , , "Extracting MZ start: " & Hex(start) & " leng: " & Hex(leng) & " ends at: " & Hex(start + leng)
        
        fp = pf & "\" & Hex(start) & ".bin"
        If fso.FileExists(fp) Then Kill fp
        
        lv2.ListItems.Add , , "Saving file as: " & fp
        
        ReDim b(leng)
        Get f, start + 1, b()
       
        f1 = FreeFile
        Open fp For Binary As f1
        Put f1, , b()
        Close f1
        
    Next
    
     Close f
    
    MsgBox "done"
    
End Sub

Function GetNext(i) As ListItem
    On Error Resume Next
 
    Set GetNext = lv.ListItems(i)
End Function

Private Sub cmdOpen_Click()
    Dim f As String
    
    f = dlg.OpenDialog(AllFiles, , , Me.hWnd)
    If Len(f) = 0 Then Exit Sub
    lv.ListItems.Clear
    lv2.ListItems.Clear
    
    LoadedFile = f
    he.LoadFile f
    
    lv2.ListItems.Add , , "Scanning " & f
    
    Dim offsets As Collection
    Dim o
    Dim li As ListItem
    Dim pe As New CPEEditor
    Dim msg As String
    
    Set offsets = FindMZOffsets(f)
    
    lv2.ListItems.Add , , "Found " & offsets.Count & " MZ markers..validating"
    
    For Each o In offsets
        
        If pe.LoadFile(f, CLng(o) + 1) Then
            msg = Hex(pe.SizeOfImage)
        Else
            msg = pe.errMessage
        End If
        
        If pe.ErrorNumber = 0 Then 'Or pe.ErrorNumber = 2 Then
            Set li = lv.ListItems.Add(, , Hex(o))
            li.Tag = CLng(o)
            'li.SubItems(1) = msg
        Else
            Set li = lv2.ListItems.Add(, , Hex(o) & " " & pe.errMessage)
            li.Tag = o
        End If
        
    Next
    
    
    
End Sub

Private Sub Form_Load()
    lv.ColumnHeaders(1).Width = lv.Width - 80
    lv2.ColumnHeaders(1).Width = lv2.Width - 80
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim offset As Long
    Dim msg As String
    
    offset = CLng(Item.Tag)
    
    msg = Item.SubItems(1)
    Me.Caption = msg
    
    he.scrollTo offset
    
    If InStr(msg, "offset:") > 0 Then
        a = InStr(msg, ":")
        a = CLng(Trim(Mid(msg, a + 1)))
        he.SelStart = offset + a
        he.SelLength = 2
    End If
    
    
    
    
End Sub

Private Sub lv2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Dim o As Long
    o = Item.Tag
    If o <> 0 Then he.scrollTo o
End Sub
