VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COllySocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'dzzie@yahoo.com
'http://sandsprite.com
Public RemoteHost As String

Private WithEvents mSck As CSocket
Attribute mSck.VB_VarHelpID = -1
Private SckSend As CSocket

Event DataReceived(data)

Property Get isUp() As Boolean
    isUp = Not (mSck Is Nothing)
End Property

Sub Listen()
    On Error Resume Next
    
    Set mSck = New CSocket
    Set SckSend = New CSocket
    
    SckSend.Protocol = sckUDPProtocol
    SckSend.RemotePort = 3333
    
    mSck.Protocol = sckUDPProtocol
    mSck.LocalPort = 4444
    mSck.Bind
    mSck.Listen
    
End Sub

Sub shutdown()
    On Error Resume Next
    mSck.CloseSocket
    Set mSck = Nothing
    Set SckSend = Nothing
End Sub

Sub SendCommand(cmd As String)
    On Error Resume Next
    SckSend.RemoteHost = RemoteHost
    SckSend.SendData cmd
    'MsgBox "Send Command: " & cmd & vbCrLf & vbCrLf & Err.Description
End Sub

Private Sub Class_Initialize()
    'RemoteHost = "10.10.10.7"
End Sub

Private Sub mSck_OnDataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim cmd As String
    mSck.GetData cmd, vbString
    RaiseEvent DataReceived(cmd)
End Sub

