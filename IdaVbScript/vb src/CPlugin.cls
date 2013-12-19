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
'dzzie@yahoo
'http://sandsprite.com

Public IDASCK As CIDASocket

Sub DoPluginAction()
    
    On Error Resume Next
    Set IDASCK = Module1.IDASCK
    Set frmPluginSample.pCplugin = Me
    'frmPluginSample.Show 1
    
    ShowWindow frmPluginSample.hWnd, 1
    
End Sub

 Sub InitPlugin()
    
    On Error Resume Next
    Set IDASCK = Module1.IDASCK
    If Not IDASCK.isUp Then IDASCK.Listen
    
 End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Dim f As Form
    For Each f In Forms
        Unload f
    Next
End Sub
