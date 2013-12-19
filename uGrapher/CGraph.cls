VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This code is based on:
'    pGRAPH.Graph.py
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


Public id
Public edges As New Collection
Public nodes As New Collection

Function nodeExists(id As Long) As Boolean
    On Error GoTo hell
    Dim c As CNode
    Set c = nodes("id:" & id)
    nodeExists = True
    Exit Function
hell:
End Function

Function FindNode(title, outNode As CNode) As Boolean
    Dim x As CNode
    For Each x In nodes
        If x.title = Trim(title) Then
            Set outNode = x
            FindNode = True
            Exit Function
        End If
    Next
End Function

Function add_node(Optional text) As CNode
    Dim n As New CNode
    n.id = nodes.Count + 1
    n.label = text
    nodes.Add n, "id:" & n.id
    Set add_node = n
End Function
 
Function add_edge(src As CNode, dst As CNode) As CEdge
    
    Dim e As New CEdge
    
    Set e.dst = dst
    Set e.src = src
    e.id = edges.Count + 1
    Set add_edge = e
    edges.Add e
    
End Function
        
    
'def del_edge(self, id = None, src = None, dst = None):
'def del_node(self, id):
    
Private Function render_graph()
    Dim udraw As String
    Dim node As CNode
    
    udraw = "["
        
    '# render each of the nodes in the graph.
    '# the individual nodes will handle their own edge rendering.
    For Each node In nodes
        udraw = udraw & node.render_node(Me) & ","
    Next

    '# trim the extraneous comment and close the graph.
    udraw = Mid(udraw, 1, Len(udraw) - 1)
    udraw = udraw & "]"
        
    render_graph = udraw
    
        
End Function

Function edges_from(node_id) As Collection

    Dim e As CEdge
    Dim c As New Collection
    
    For Each e In edges
        If e.src.id = node_id Then
            c.Add e
        End If
    Next
    
    Set edges_from = c
    
End Function


Function GenerateGraph()
    Dim command As String
    
    command = "graph(new_placed(" & _
                    render_graph() & _
                   "))"
                   
    frmMain.udraw.SendCommand command


End Function

 
