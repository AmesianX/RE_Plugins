VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This code is based on:
'    pGRAPH.Node.py
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
Public Number
Public label
Public color
Public shape
Public title
Public va As Long
Public Tag As Variant

Sub SetColor(m_color)
    Dim command As String
    
    color = m_color
    
    command = "graph(change_attr([" & _
                   "node(""" & id & """,[a(""COLOR"",""" & color & """)])" & _
                   "]))"
                   
                   
    frmMain.udraw.SendCommand command
    
End Sub

Sub SelectChildren()
    Dim command As String
    
    command = "menu(navigation(select_siblings([""" & id & """])))"
    frmMain.udraw.SendCommand command
    
End Sub
Sub SetFocus(Optional animated = True):
        '''
        'Scrolls the visible part of the graph visualization to the node specified by "node_id".

'        @todo: This routine is buggy. Appears to only want to work when being called after a call to
 '              change_element_color(), though the element color change will not actually work. Need to debug.
        '''

        Dim command As String
        
        command = "special(focus_node_animated(""" & id & """))"
        
        If Not animated Then command = Replace(command, "_animated", Empty)

        frmMain.udraw.SendCommand command

End Sub


Private Sub Class_Initialize()
    id = 0
    Number = 0
    color = "white"
    label = ""
    shape = "box"
End Sub

Private Function GetMenus()
    Dim r As String
    Dim n As CNodeMenu
    
    If frmMain.udraw.NodeMenus.Count = 0 Then Exit Function
    
    r = "m(["
    For Each n In frmMain.udraw.NodeMenus
        r = r & "menu_entry(""" & n.id & """,""" & n.label & """),"
    Next
    
    r = Mid(r, 1, Len(r) - 1)
    r = r & "]),"
    
    GetMenus = r
    
End Function
Function render_node(graph As CGraph)
'        Render a node description suitable for use in a uDraw file using the set internal attributes.
'
        Dim udraw As String

       ' udraw  = 'l("%08x",'                            % self.id
       ' udraw +=   'n("",'                              # open node
       ' udraw +=     '['                                # open attributes
       ' udraw +=       'a("_GO","%s"),'                 % self.shape
       ' udraw +=       'a("COLOR","#%06x"),'            % self.color
       ' udraw +=       'a("OBJECT","%s"),'              % self.label
       ' udraw +=       'a("FONTFAMILY","courier"),'
       ' udraw +=       'a("BORDER","none")'
       ' udraw +=     '],'                               # close attributes
       ' udraw +=     '['                                # open edges

         udraw = "l(""" & id & """," & _
                 "n(""""," & _
                 "[" & _
                 "a(""OBJECT"",""" & label & """)," & _
                 GetMenus() & _
                 "],["

        '"m([menu_entry(""2"",""test"")])," & _

        Dim edges As Collection
        Dim edge As CEdge
        
        Set edges = graph.edges_from(id)

        For Each edge In edges:
            udraw = udraw & edge.render_edge_udraw() & ","
        Next
        
        '# trim the extraneous comment and close the graph.
        If edges.Count > 0 Then udraw = Mid(udraw, 1, Len(udraw) - 1)
        udraw = udraw & "]))"

        render_node = udraw

End Function



'    ####################################################################################################################
'    def render_node_udraw_update(self):
'        '''
'        Render a node update description suitable for use in a uDraw file using the set internal attributes.
'
'        @rtype:  String
'        @return: uDraw node update description.
'        ''''
'
'        udraw  = 'new_node("%08x","",'                % self.id
'        udraw +=   '['
'        udraw +=     'a("_GO","%s"),'                 % self.shape
'        udraw +=     'a("COLOR","#%06x"),'            % self.color
'        udraw +=     'a("OBJECT","%s"),'              % self.label
'        udraw +=     'a("FONTFAMILY","courier"),'
'        udraw +=     'a("BORDER","none")'
'        udraw +=   ']'
'        udraw += ')'
'
'        return udraw


