Attribute VB_Name = "Module1"
Private Type COPYDATASTRUCT
    dwFlag As Long
    cbSize As Long
    lpData As Long
End Type

 Public Const GWL_WNDPROC = (-4)
 Public Const WM_COPYDATA = &H4A
 Global lpPrevWndProc As Long
 Global gHW As Long
 Global IDA_HWND As Long
 Global ResponseBuffer As String
 
 Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
 Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

 Public Sub Hook(hwnd As Long)
     gHW = hwnd
     lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
     'Debug.Print lpPrevWndProc
 End Sub

 Public Sub Unhook()
     Dim Temp As Long
     Temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
 End Sub

 Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
     If uMsg = WM_COPYDATA Then RecieveTextMessage lParam
     WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
 End Function

Private Sub RecieveTextMessage(lParam As Long)
   
    Dim CopyData As COPYDATASTRUCT
    Dim Buffer(1 To 2048) As Byte
    Dim Temp As String
    Dim lpData As Long
    Dim sz As Long
    Dim tmp() As Byte
    ReDim tmp(30)
    
    CopyMemory CopyData, ByVal lParam, Len(CopyData)
    
    If CopyData.dwFlag = 3 Then
    
        CopyMemory tmp(0), ByVal lParam, Len(CopyData)
        'Text1 = HexDump(tmp, Len(CopyData))
        
        lpData = CopyData.lpData
        sz = CopyData.cbSize
        
        CopyMemory Buffer(1), ByVal lpData, sz
        Temp = StrConv(Buffer, vbUnicode)
        Temp = Left$(Temp, InStr(1, Temp, Chr$(0)) - 1)
        'heres where we work with the intercepted message
        Form1.List2.AddItem "Recv(" & Temp & ")"
        Form1.List2.AddItem ""
        ResponseBuffer = Temp
    End If
     
End Sub

Sub SendCMD(msg As String)
    Dim cds As COPYDATASTRUCT
    Dim ThWnd As Long
    Dim buf(1 To 255) As Byte
    
    ResponseBuffer = Empty
    Form1.List2.AddItem "SendingCMD(hwnd=" & IDA_HWND & ", msg=" & msg & ")"
    
    Call CopyMemory(buf(1), ByVal msg, Len(msg))
    cds.dwFlag = 3
    cds.cbSize = Len(msg) + 1
    cds.lpData = VarPtr(buf(1))
    i = SendMessage(IDA_HWND, WM_COPYDATA, gHW, cds)
    'since SendMessage is syncrnous if the command has a response it will be received before this returns..
    
End Sub

Function SendCmdRecvText(cmd As String) As String
    SendCMD cmd
    SendCmdRecvText = ResponseBuffer
End Function

Function SendCmdRecvLong(cmd As String) As Long
    SendCMD cmd
    On Error Resume Next
    SendCmdRecvLong = CLng(ResponseBuffer)
End Function

