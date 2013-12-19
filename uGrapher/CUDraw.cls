VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CuDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This code is based on:
'    uDraw Connector
'    Copyright (C) 2006 Pedram Amini <pedram.amini@gmail.com>
'    contact:      pedram.amini@gmail.com
'    organization: www.openrce.org
'
'Ported to vb by: dzzie@yahoo.com
'
'License: GPL
'
' This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public
' License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later
' version.
'
' This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
' warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License along with this program; if not, write to the Free
' Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA




Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long



Public WithEvents ws As Winsock
Attribute ws.VB_VarHelpID = -1
Public WithEvents tmr As Timer
Attribute tmr.VB_VarHelpID = -1

Public uDrawPath As String
Public graph As New CGraph
Public ParentHWND As Long
Public connected As Boolean

Dim m_hwnd As Long
Dim command As String
Dim timeout As Boolean
Dim selNode As CNode
Dim pid As Long

Public NodeMenus As New Collection

Event Message(msg)
Event NodeSelected(n As CNode)
Event DoubleClick(n As CNode)
Event PopupMenuSel(n As CNode, sMenu As String)


Sub AddNodeMenu(label As String)
    Dim n As New CNodeMenu
    n.id = NodeMenus.Count + 1
    n.label = label
    NodeMenus.Add n, "id:" & n.id
End Sub

Private Function GetMenuForID(id As Long) As CNodeMenu
    On Error Resume Next
    Set GetMenuForID = NodeMenus("id:" & id)
End Function

Private Function MenuExistsForID(id As Long) As Boolean
    If Not GetMenuForID(id) Is Nothing Then MenuExistsForID = True
End Function

Public Property Get hwnd()
    
   ' Me.SetTitle "uDraw(Graph)"
    If m_hwnd = 0 Then
        m_hwnd = FindWindow("TkTopLevel", "uDraw(Graph) 3.1.1")
    End If
    
    hwnd = m_hwnd
    
    

End Property

Sub ShutDown()
    SendCommand "menu(file(exit))"
End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)
    On Error Resume Next
    ws.Accept requestID
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    Dim s As String
    Dim a, b, id, i, j
    Dim tmp() As String
    Dim y() As String
    
    On Error GoTo hell
    
    ws.GetData s
    
    tmp = Split(s, vbCrLf)
    For i = 0 To UBound(tmp)
        s = tmp(i)
        If Len(s) > 0 Then
        
            'If ParentHWND > 0 Then SetParent hwnd, ParentHWND
            RaiseEvent Message(s)
            
            Const selections = "node_selections_labels"
            If Left(s, Len(selections)) = selections Then
                s = Replace(s, selections, Empty)
                s = Replace(s, "(", Empty)
                s = Replace(s, ")", Empty)
                s = Replace(s, "[", Empty)
                s = Replace(s, "]", Empty)
                s = Replace(s, """", Empty)
                y = Split(s, ",")
                For j = 0 To UBound(y)
                    If Len(y(i)) > 0 Then
                        If graph.nodeExists(CLng(y(j))) Then
                            Set selNode = graph.nodes("id:" & CLng(y(j)))
                            RaiseEvent NodeSelected(selNode)
                        End If
                    End If
                Next
            End If
            
            Const popup = "popup_selection_node"
            If Left(s, Len(popup)) = popup Then
                s = Replace(s, popup, Empty)
                s = Replace(s, "(", Empty)
                s = Replace(s, ")", Empty)
                s = Replace(s, """", Empty)
                y = Split(s, ",")
                If graph.nodeExists(CLng(y(0))) Then
                    If MenuExistsForID(CLng(y(1))) Then
                        Set selNode = graph.nodes("id:" & CLng(y(0)))
                        RaiseEvent PopupMenuSel(selNode, GetMenuForID(CLng(y(1))).label)
                    End If
                End If
            End If
                    
                
            If InStr(s, "node_double_click") > 0 Then
                If Not selNode Is Nothing Then
                    RaiseEvent DoubleClick(selNode)
                End If
            End If
            
        End If
    Next
    
hell:
End Sub



Function Initilize(w As Winsock, t As Timer, Optional upth As String) As Boolean

    If Len(upth) > 0 Then
        uDrawPath = upth
    Else
        If Len(uDrawPath) = 0 Then
           uDrawPath = "C:\Program Files\uDraw(Graph)\bin\uDrawGraph.exe"
        End If
    End If
    
    If Dir(uDrawPath) = "" Then
        MsgBox "Could not locate uDrawGraph.exe do you have it installed?", vbInformation
    End If
    
    Set ws = w
    Set tmr = t
    
    On Error GoTo hell
    
    If Not Connect Then
         pid = Shell(uDrawPath & " -server")
         If Not Connect Then
            Initilize = False
            Exit Function
         End If
    End If
    
    connected = True
    Initilize = True
    
    Exit Function
hell:
    
End Function

Private Function Connect() As Boolean
    On Error Resume Next
    
    tmr.Interval = 1000
    ws.Close
    ws.Connect "localhost", 2542
    
    timeout = False
    tmr.Enabled = True
    Do While tmr.Enabled
        DoEvents
        If ws.State = sckClosed Then Exit Do
        If ws.State = sckConnected Then
            tmr.Enabled = False
            Connect = True
            Exit Function
        End If
    Loop
    
End Function



Sub SendCommand(cmd As String)
    On Error Resume Next
    
    If Not connected Then
        MsgBox "not connected!"
    End If
    
    ws.SendData cmd & vbCrLf
    
End Sub



Sub layout_improve_all()

        ''
        'This command starts the layout algorithm to improve the visualization quality of the whole graph by reducing
        'unnecessary edge crossings and edge bends.
        ''

        command = "menu(layout(improve_all))" & vbCrLf
        ws.SendData (command)

End Sub

'Sub open_survey_view()
'        'Open a survey view showing the whole graph in a reduced scale.
'        ws.SendData "menu(view(open_survey_view))" & vbCrLf
'End Sub
   

Sub scaleit(p)
        'Sets the scale to the given parameter which is a percent value that must be from 1 to 100.
        
        If p = "full_scale" Or p = "full" Then
           p = "full_scale"
        ElseIf p = "fit" Then
           p = "fit_scale_to_window"
        ElseIf IsNumeric(p) Then
           p = "scale(" & p & ")"
        End If
        
        ws.SendData "menu(view(" & p & "))" & vbCrLf
        
End Sub

Sub SetTitle(msg)
        ws.SendData "window(title(""" & msg & """))"
End Sub


Private Sub tmr_Timer()
    tmr.Enabled = False
End Sub
