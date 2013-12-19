VERSION 5.00
Begin VB.Form frmManualPatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Hex Bytes Using Space As Dividers"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5010
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkPatchFile 
      Caption         =   "Actually Patch EXE along with IDB (No Backups made)"
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   4815
   End
   Begin VB.TextBox txtLeng 
      Height          =   285
      Left            =   2820
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtCurrent 
      BackColor       =   &H80000004&
      Height          =   315
      Left            =   660
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   420
      Width           =   4155
   End
   Begin VB.CommandButton cmdGetBytes 
      Caption         =   "Display Bytes"
      Height          =   315
      Left            =   3420
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.TextBox txtOffset 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdPatch 
      Caption         =   "Patch"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3780
      TabIndex        =   1
      Top             =   3300
      Width           =   1215
   End
   Begin VB.TextBox txtBytes 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   3300
      Width           =   3615
   End
   Begin VB.Label lblFileOffset 
      Height          =   315
      Left            =   900
      TabIndex        =   11
      Top             =   2460
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "FileOffset"
      Height          =   255
      Left            =   60
      TabIndex        =   10
      Top             =   2460
      Width           =   1035
   End
   Begin VB.Label lblAsm 
      Caption         =   "Asm :"
      Height          =   1395
      Left            =   60
      TabIndex        =   8
      Top             =   900
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "Bytes"
      Height          =   315
      Left            =   60
      TabIndex        =   7
      Top             =   420
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Virtual Address                            Len"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   2715
   End
End
Attribute VB_Name = "frmManualPatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fileOffset As Long
Dim curFile As String

Private Sub cmdGetBytes_Click()
    Dim o As Long
    Dim rva As Long
    
    If Not GetHextxt(txtOffset, o) Then Exit Sub
    
    rva = o - clsSect.ImageBase
    
    txtCurrent = HexDumpBytes(o, txtLeng)
    lblAsm.Caption = GetAsmRange(o, txtLeng)
    
    fileOffset = clsSect.RvaToOffset(rva)
    
    lblFileOffset.Caption = Hex(fileOffset)
    
    cmdPatch.enabled = True
    
    
End Sub

Private Sub cmdPatch_Click()

    Dim o As Long, i As Long, x As Long
    Dim tmp() As String
    Dim f As Long
    
    On Error GoTo hell
    
    If chkPatchFile.Value = 1 Then
        f = FreeFile
        Open curFile For Binary As f
    End If
    
    GetHextxt txtOffset, o
    
    tmp = Split(txtBytes, " ")
    
    For i = 0 To UBound(tmp)
        If tmp(i) <> "" Then
            x = CLng("&h" & tmp(i))
            PatchByte (o + i), CByte(x)
            If chkPatchFile.Value = 1 Then
                Put f, (fileOffset + i + 1), CByte(x)
            End If
        End If
    Next
    
    If f > 0 Then Close f
    
MsgBox "Patch complete"

Exit Sub
hell:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    
    
    curFile = loadedFile
    txtOffset = ScreenEA
    If curFile <> clsSect.curFile Then
        clsSect.LoadSections curFile
    End If
    
End Sub

Private Sub txtOffset_Change()
 cmdPatch.enabled = False
End Sub
