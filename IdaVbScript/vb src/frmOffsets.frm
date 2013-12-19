VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOffsets 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdJump 
      Caption         =   "Jump"
      Height          =   375
      Left            =   2670
      TabIndex        =   11
      Top             =   1470
      Width           =   945
   End
   Begin VB.TextBox txtBytes 
      BackColor       =   &H80000000&
      Height          =   315
      Left            =   3660
      TabIndex        =   10
      Top             =   2280
      Width           =   2715
   End
   Begin VB.TextBox txtOffset 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1440
      TabIndex        =   8
      Top             =   2220
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "FileOffset"
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "RVA"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "VirtAddress"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   1440
      Width           =   1155
   End
   Begin VB.TextBox txtRVA 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtVA 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvSect 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   2355
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
         Text            =   "Name"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Virtual Addr"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Virtual Size"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "RawOffset"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "RawSize"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Attributes"
         Object.Width           =   1235
      EndProperty
   End
   Begin VB.Label lblSection 
      Height          =   255
      Left            =   3660
      TabIndex        =   9
      Top             =   1920
      Width           =   2715
   End
   Begin VB.Label Label1 
      Caption         =   "Section:             Bytes :"
      Height          =   555
      Left            =   2820
      TabIndex        =   1
      Top             =   1920
      Width           =   735
   End
End
Attribute VB_Name = "frmOffsets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private selIndex As Long

Private Sub cmdJump_Click()
    On Error Resume Next
    Jump CLng("&h" & txtVA)
End Sub

Private Sub Form_Load()
    Dim curFile As String
    
    selIndex = 1
    curFile = loadedFile
    
    If curFile <> clsSect.curFile Then
        clsSect.LoadSections curFile
    End If
    
    clsSect.FilloutListView lvSect
    Me.Caption = "ImageBase: " & Hex(clsSect.ImageBase)
    
End Sub


Private Sub cmdCalculate_Click()
    Dim va As Long
    Dim fo As Long
    Dim rva As Long
    Dim sectName As String
    
    On Error Resume Next
    
    Select Case selIndex
        Case 0:  'virtual address
                If Not GetHextxt(txtVA, va) Then Exit Sub
                
                If va < clsSect.ImageBase Then
                    MsgBox "VA is below Imagebase"
                    Exit Sub
                End If
                
                rva = va - clsSect.ImageBase
                fo = clsSect.RvaToOffset(rva, , sectName)
                
                txtRVA = Hex(rva)
                txtOffset = Hex(fo)
        Case 1: 'rva
                If Not GetHextxt(txtRVA, rva) Then Exit Sub
                
                va = rva + clsSect.ImageBase
                fo = clsSect.RvaToOffset(rva, , sectName)
                
                txtVA = Hex(va)
                txtOffset = Hex(fo)
        Case 2: 'file offset
                If Not GetHextxt(txtOffset, fo) Then Exit Sub
                
                If fo > FileLen(clsSect.curFile) Then
                    MsgBox "This file offset does not exist in loaded file!"
                    Exit Sub
                End If
                
                rva = clsSect.OffsetToRVA(fo, sectName)
                va = rva + clsSect.ImageBase
              
                txtRVA = Hex(rva)
                txtVA = Hex(va)
    End Select
        
    lblSection.Caption = sectName
    txtBytes = HexDumpBytes(va, 8)
    
End Sub


Private Sub Option1_Click(Index As Integer)

    Enable txtVA, False
    Enable txtRVA, False
    Enable txtOffset, False
    
    Select Case Index
        Case 0: Enable txtVA
        Case 1: Enable txtRVA
        Case 2: Enable txtOffset
    End Select
        
    selIndex = Index
End Sub

