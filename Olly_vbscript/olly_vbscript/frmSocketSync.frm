VERSION 5.00
Begin VB.Form frmSocketSync 
   Caption         =   "IDA Olly Socket Sync"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Remote IDB"
      Height          =   285
      Left            =   2835
      TabIndex        =   10
      Top             =   1395
      Width           =   1275
   End
   Begin VB.CheckBox chkDebugMode 
      Caption         =   "Capture Debug Log"
      Height          =   285
      Left            =   45
      TabIndex        =   9
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   330
      Left            =   2475
      TabIndex        =   8
      Top             =   1845
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View"
      Height          =   330
      Left            =   3510
      TabIndex        =   7
      Top             =   1845
      Width           =   645
   End
   Begin VB.TextBox txtModule 
      Height          =   330
      Left            =   1755
      TabIndex        =   5
      Top             =   495
      Width           =   2310
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Save Setting"
      Height          =   315
      Left            =   2835
      TabIndex        =   3
      Top             =   945
      Width           =   1290
   End
   Begin VB.CheckBox chkSync 
      Caption         =   "Enable Remote Syncing"
      Height          =   375
      Left            =   45
      TabIndex        =   2
      Top             =   945
      Width           =   2115
   End
   Begin VB.TextBox txtRemoteHost 
      Height          =   315
      Left            =   1740
      TabIndex        =   1
      Top             =   60
      Width           =   2310
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status: Closed"
      Height          =   330
      Left            =   45
      TabIndex        =   6
      Top             =   1440
      Width           =   1635
   End
   Begin VB.Label Label3 
      Caption         =   "Mdule To Sync"
      Height          =   240
      Left            =   45
      TabIndex        =   4
      Top             =   540
      Width           =   1725
   End
   Begin VB.Label Label1 
      Caption         =   "Remote Host to Sync"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmSocketSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'dzzie@yahoo.com
'http://sandsprite.com

Public pPlugin As CPlugin

Private Sub chkDebugMode_Click()
    pPlugin.DebugMode = IIf(chkDebugMode.Value = 1, True, False)
End Sub

Private Sub cmdDone_Click()

        pPlugin.OlySck.RemoteHost = txtRemoteHost.Text
        pPlugin.ModuleToSync = txtModule
        
        SetSyncFlag chkSync.Value
        
        If chkSync.Value = 1 Then
            If Not pPlugin.OlySck.isUp() Then
                pPlugin.OlySck.Listen
            End If
        Else
            pPlugin.OlySck.shutdown
        End If
        
        lblStatus = IIf(pPlugin.OlySck.isUp, "Listening", "Closed")
        
        'Unload Me
        
End Sub

Private Sub Command1_Click()
    
    frmOllyScript.txtScript = pPlugin.DebugLog
    'frmOllyScript.txtScript.Move 0, 0, frmOllyScript.Width, frmOllyScript.Height
    frmOllyScript.Show 1

End Sub

Private Sub Command2_Click()
    pPlugin.DebugLog = Empty
End Sub

Private Sub Command3_Click()
    If pPlugin.OlySck.isUp Then
        pPlugin.OlySck.SendCommand "curidb"
    Else
        MsgBox "You have to enable the socket and save settings first...", vbInformation
    End If
End Sub

Private Sub Form_Load()
    
    If pPlugin.OlySck.isUp Then
        chkSync.Value = 1
        lblStatus = "Listening"
    End If
    
    chkDebugMode.Value = IIf(pPlugin.DebugMode, 1, 0)
    pPlugin.OlySck.RemoteHost = GetSetting("OllySync", "Settings", "RemoteHost", "127.0.0.1")
    pPlugin.ModuleToSync = GetSetting("OllySync", "Settings", "ModuleToSync", "test.dll")
    
    txtRemoteHost = pPlugin.OlySck.RemoteHost
    txtModule.Text = pPlugin.ModuleToSync
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "OllySync", "Settings", "RemoteHost", txtRemoteHost.Text
    SaveSetting "OllySync", "Settings", "ModuleToSync", txtModule.Text
End Sub
