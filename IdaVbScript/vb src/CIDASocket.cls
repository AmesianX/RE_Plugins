VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIDASocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents mSck As CSocket
Attribute mSck.VB_VarHelpID = -1
Private WithEvents sckSend As CSocket
Attribute sckSend.VB_VarHelpID = -1
Private WithEvents ollySck As CSocket
Attribute ollySck.VB_VarHelpID = -1

Const DONE_MSG = "ids:done"

Property Get isUp() As Boolean
    isUp = Not (mSck Is Nothing)
End Property

Sub Listen()
    On Error Resume Next
    Set mSck = New CSocket
    Set sckSend = New CSocket
    Set ollySck = New CSocket
    
    ollySck.Protocol = sckUDPProtocol
    ollySck.LocalPort = 3333
    ollySck.Bind
    ollySck.Listen
    
    mSck.Protocol = sckTCPProtocol
    mSck.LocalPort = 2222
    mSck.Listen
    
End Sub

Sub shutdown()
    On Error Resume Next
    mSck.CloseSocket
    Set mSck = Nothing
End Sub
 

Private Sub mSck_OnConnectionRequest(ByVal requestID As Long)
    On Error Resume Next
    sckSend.Accept requestID
    MessageUI "IDASocket Accepted " & requestID & " from " & mSck.RemoteHostIP & " " & Err.Description & vbCrLf
End Sub

Private Sub ollySck_OnDataArrival(ByVal bytesTotal As Long)
    
    Dim cmd As String
    Dim args() As String
    On Error Resume Next
    
    ollySck.GetData cmd, vbString

    args = Split(cmd, " ")
    
    Select Case args(0)
        Case "jmp": Jump CLng(args(1))
        Case "jmpfunc": Jump GetFuncOffset(args(1))
    End Select
    
End Sub

Private Sub sckSend_OnDataArrival(ByVal bytesTotal As Long)
    On Error GoTo hell
    
    Dim cmd As String
    Dim args() As String
    Dim resp
    
    sckSend.GetData cmd, vbString
    
    MessageUI "Received Cmd: " & cmd & vbCrLf
    
    args = Split(cmd, " ")
    
    
    Select Case args(0)
    
        Case "jmp": Jump CLng(args(1)) 'this is for olly integration command set
                    Exit Sub
                    
        Case "jmpfunc":
                        MessageUI "Function offset for " & args(1) & " = " & GetFuncOffset(args(1))
                        Jump GetFuncOffset(args(1))
                        Exit Sub
                        
        Case "getfunc":
                        MessageUI "getfunc(" & args(1) & ") = " & GetFuncOffset(args(1))
                        sckSend.SendData GetFuncOffset(args(1)) & vbCrLf & DONE_MSG
                        Exit Sub
                        
        Case "newname":
                        MessageUI "newname( " & args(1) & " , " & args(2) & " ) = " & GetFuncOffset(args(1))
                        cIda.Setname GetFuncOffset(args(1)), args(2)
                        sckSend.SendData DONE_MSG
                        Exit Sub
                        
        Case "ping":
                        MessageUI "PING"
                        sckSend.SendData "PONG" & vbCrLf & DONE_MSG
                        Exit Sub

                         
    End Select
        
    Select Case UBound(args)
        Case 0:  resp = CallByName(cIda, CStr(args(0)), VbMethod)
        Case 1:  resp = CallByName(cIda, CStr(args(0)), VbMethod, args(1))
        Case 2:  resp = CallByName(cIda, CStr(args(0)), VbMethod, args(1), args(2))
        Case 3:  resp = CallByName(cIda, CStr(args(0)), VbMethod, args(1), args(2), args(3))
    End Select
    
1    If IsArray(resp) Then
        'If aryistmpty(resp(0)) Then
        '
        'If TypeName(resp(0)) = "Byte" Then
        '    sckSend.SendData resp 'StrConv(resp, vbUnicode)
        'Else
4            sckSend.SendData Join(resp, vbCrLf)
        'End If
5    ElseIf Len(resp) > 0 Then
6        sckSend.SendData resp
    End If
    
    'MessageUI "Send Done Msg"
    sckSend.SendData DONE_MSG
    
Exit Sub
hell:
        sckSend.SendData Err.Number & ":" & Err.Description & "Line: " & Erl
        sckSend.SendData DONE_MSG
        
End Sub










'        Case "patch": PatchByte CLng(args(1)), CByte(args(2))
'        Case "undef": Undefine CLng(args(1))
'        Case "tocode": MakeCode CLng(args(1))
'        Case "name": Setname CLng(args(1)), args(2)
'        Case "unname":
'        Case "refresh":  aRefresh
'        Case "comment": cIda.AddComment args(1), COLOR_DEFAULT
'
'        Case "orgbyte": resp = OriginalByte(CLng(args(1)))
'                        mSck.SendData resp
'
'        Case "getasm":  resp = GetAsmCode(CLng(args(1)))
'                        mSck.SendData resp
'
'        Case "getbytes": resp = cIda.GetBytes(args(1), args(2))
'                         mSck.SendData StrConv(resp, vbUnicode)
'
'        Case "fpath": resp = loadedFile
'                      mSck.SendData resp
          
          
  
