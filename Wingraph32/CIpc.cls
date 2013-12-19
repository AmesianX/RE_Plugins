VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents sc As clsSubClass
Attribute sc.VB_VarHelpID = -1
Private WithEvents tmr As CTimer
Attribute tmr.VB_VarHelpID = -1

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Type COPYDATASTRUCT
    dwFlag As Long
    cbSize As Long
    lpData As Long
End Type

Private Const WM_COPYDATA = &H4A

'Public BlockingMode As Boolean
Public timeout As Long
Private DataIn As String

Event DataReceived(msg As String) 'only raised when not in blocking mode
Event SendTimedOut()              'only sent for first timeout until TimedOut var reset to false
Event DataSend(msg As String, isBlocking As Boolean)
Event Error(msg As String)

Private ServerHwnd As Long
Private myWindow As Long
Public ServerName As String
Public ClientHWND As Long
Private WaitForDone As Boolean 'receives up to DONE message or timeout

Private isListening As Boolean
Public TimedOut As Boolean 'consumer is responsible for setting this to false when they see fit to retest

Property Get IamListening() As Boolean
    IamListening = isListening
End Property

Function Listen(Optional name As String, Optional hwnd As Long) As String

    If isListening Then Exit Function
    
    If hwnd = 0 Then
        myWindow = True
        ServerHwnd = CreateWindowEx(0&, "STATIC", "MESSAGE_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)
        myWindow = ServerHwnd
    Else
        ServerHwnd = hwnd
        myWindow = 0
    End If
    
    If Len(name) > 0 Then ServerName = name
    If Len(ServerName) = 0 Then Err.Raise 0, "CIPC::Listen", "Invalid ServerName!"
    
    If sc.AttachMessage(ServerHwnd, WM_COPYDATA) Then
        SaveSetting "IPC", "Handles", ServerName, ServerHwnd
        isListening = True
    Else
        Err.Raise 1, "CIPC::Listen", "Failed to subclass IPC window!"
        isListening = False
    End If
        
End Function

Function ShutDown(Optional name As String)
    Dim hwnd As Long
    
    If Len(name) = 0 Then
        hwnd = GetSetting("IPC", "Handles", ServerName)
    Else
        hwnd = ServerHwnd
    End If
    
    sc.DetatchMessage ServerHwnd, WM_COPYDATA
    SaveSetting "IPC", "Handles", ServerName, 0
    
    If myWindow = hwnd Then
        DestroyWindow myWindow
        myWindow = 0
    End If
        
End Function

'blocking call
Function SendAndRecv(msg As String) As String

    If ClientHWND = 0 Then Err.Raise 0, , "Call FindClient First"
    If Len(ServerName) = 0 Then Err.Raise 0, , "Set ServerName First"
    If ServerHwnd = 0 Then Err.Raise 0, , "Set ServerHwnd First"
    
    WaitForDone = False
    InternalSend msg
    WaitLoop
    SendAndRecv = DataIn
    
End Function

'blocking call waits till DONE message is received from remote client
Function SendAndRecvTillDone(msg As String) As String

    If ClientHWND = 0 Then Err.Raise 0, , "Call FindClient First"
    If Len(ServerName) = 0 Then Err.Raise 0, , "Set ServerName First"
    If ServerHwnd = 0 Then Err.Raise 0, , "Set ServerHwnd First"
    
    WaitForDone = True
    InternalSend msg
    WaitLoop
    SendAndRecvTillDone = DataIn
    
End Function


Function FindClient(client As String) As Boolean
    Dim hwnd As Long
    On Error Resume Next
    hwnd = CLng(GetSetting("IPC", "Handles", client, 0))
    If hwnd <> 0 Then
        If IsWindow(hwnd) = 1 Then
            FindClient = True
            ClientHWND = hwnd
        Else
            SaveSetting "IPC", "Handles", client, 0
            ClientHWND = 0
            FindClient = False
        End If
    End If
End Function

Private Function InternalSend(msg As String) As Boolean
    On Error Resume Next
    
    Dim Buffer(1 To 2048) As Byte
    Dim CopyData As COPYDATASTRUCT
    
    If IsWindow(ClientHWND) = 0 Then
        RaiseEvent Error("Hwnd: " & Hex(ClientHWND) & " does not exist")
        InternalSend = False
        Exit Function
    End If
    
    RaiseEvent DataSend(msg, False)
    
    DataIn = Empty
    CopyMemory Buffer(1), ByVal msg, Len(msg)
    CopyData.dwFlag = 3
    CopyData.cbSize = Len(msg) + 1
    CopyData.lpData = VarPtr(Buffer(1))
    SendMessage ClientHWND, WM_COPYDATA, ClientHWND, CopyData
    InternalSend = True
    
End Function

Function Send(msg As String) As Boolean
    
    If ClientHWND = 0 Then Err.Raise 0, , "Call findClient First"
    
    WaitForDone = False
    Send = InternalSend(msg)

End Function

Private Sub Class_Initialize()
    Set sc = New clsSubClass
    Set tmr = New CTimer
End Sub

Private Sub Class_Terminate()
    ShutDown
    SaveSetting "IPC", "Handles", ServerName, 0
    If myWindow <> 0 Then DestroyWindow ServerHwnd
End Sub

Private Sub sc_MessageReceived(hwnd As Long, wMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
      Dim msg As String
    
      If wMsg = WM_COPYDATA Then
            If RecieveTextMessage(lParam, msg) Then
                tmr.Reset
                If msg = "DONE" Then
                    WaitForDone = False
                    msg = Empty
                End If
                If WaitForDone Then
                    DataIn = DataIn & msg
                Else
                    DataIn = DataIn & msg
                    RaiseEvent DataReceived(DataIn)
                    tmr.Disable
                End If
            End If
      End If
            
End Sub

Private Function RecieveTextMessage(lParam As Long, msg As String) As Boolean
   
    Dim CopyData As COPYDATASTRUCT
    Dim Buffer(1 To 2048) As Byte
    Dim Temp As String
    
    msg = Empty
    
    CopyMemory CopyData, ByVal lParam, Len(CopyData)
    
    If CopyData.dwFlag = 3 Then
        CopyMemory Buffer(1), ByVal CopyData.lpData, CopyData.cbSize
        Temp = StrConv(Buffer, vbUnicode)
        Temp = Left$(Temp, InStr(1, Temp, Chr$(0)) - 1)
        'heres where we work with the intercepted message
        msg = Temp
        RecieveTextMessage = True
    End If
    
End Function

Private Sub WaitLoop()
    If timeout = 0 Then timeout = 3000
    tmr.EnableTimer timeout
    Do While Len(DataIn) = 0
        DoEvents
        If Not tmr.Enabled Then
            If Not TimedOut Then RaiseEvent SendTimedOut
            TimedOut = True
            Exit Do
        End If
    Loop
    tmr.Disable
End Sub

Private Sub tmr_Timer()
    tmr.Disable
End Sub
