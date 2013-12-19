VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'dzzie@yahoo.com
'http://sandsprite.com
Public WithEvents OlySck As COllySocket
Attribute OlySck.VB_VarHelpID = -1

Public ModuleToSync As String
Public DebugLog As String
Public DebugMode As Boolean

Public Sub AppendToDebugLog(Msg)
    If Not DebugMode Then Exit Sub
    DebugLog = DebugLog & vbCrLf & Msg
End Sub

Private Sub Class_Initialize()
    Set OlySck = New COllySocket
    DebugMode = False
End Sub

Sub DoPluginAction(action As Long)
    
    AppendToDebugLog "DoPluginAction received command: " & action
    
    Select Case action
        Case 0:
                Bpx_Handler = 1
                frmInstance.Show 1
        Case 1:
                Set frmSocketSync.pPlugin = Me
                frmSocketSync.Show 1
        Case Else
            MsgBox "Unknown Action Index: " & action
    End Select
    
End Sub

Sub BreakPoint(address As Long)
    
    AppendToDebugLog "BreakPoint at address: " & address
    frmInstance.HandleBreakpoint
    
End Sub

Sub FireSync(eventid As Long, arg As Long)
    
    On Error GoTo hell
    
    Dim va As Long
    Dim rva As Long
    Dim imageBase As Long
    Dim module As String
    
    'If Not OlySck.isUp Then Exit Sub
    
    AppendToDebugLog "FireSync eventid: " & eventid & " argument: " & arg
    
    Select Case eventid
        Case 0: 'breakpoint or single step
                va = arg
                If GetImageBaseAndNameForVA(va, imageBase, module) Then
                    rva = va - imageBase
                    If Me.ModuleToSync = "" Or InStr(1, Me.ModuleToSync, module, vbTextCompare) > 0 Then
                        OlySck.SendCommand "jmp_rva " & rva 'rva so it works with dlls or exes...
                        AppendToDebugLog "FireSync target module Found! sending jmprva " & rva
                    Else
                        AppendToDebugLog "did not match target mod " & Me.ModuleToSync & " name was: " & module
                    End If
                Else
                    AppendToDebugLog "GetImageBaseAndNameForVa Failed: va: " & va
                End If
    End Select
    
    Exit Sub
hell:
     AppendToDebugLog "Error in FireSync: Line:" & Erl() & " Desc: " & Err.Description
    'MsgBox eventid & " " & arg
End Sub
 

Private Sub OlySck_DataReceived(data As Variant)
    On Error Resume Next
    Dim f As Form
    Dim ff As frmSocketSync
    Dim tmp
    Dim a
    
    'MsgBox "OlySck_DataReceived:" & data
    AppendToDebugLog "OllySck data received: " & data
    
    tmp = Split(data, " ")
    Select Case tmp(0)
    
        Case "curidb":
            data = Replace(data, "curidb ", Empty)
            a = InStrRev(data, "\")
            If a > 0 Then
                MsgBox Mid(data, a)
            Else
                MsgBox data
            End If
            
    End Select
            
            
End Sub
