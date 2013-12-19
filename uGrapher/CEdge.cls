VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEdge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This code is based on:
'    pGRAPH.CEdge
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
Public src As CNode
Public dst As CNode
Public color
Public label

Sub class_initilize()
    color = 0
End Sub

Function render_edge_udraw():

        '# ensure nodes exist at the source and destination of this edge.
        If src Is Nothing Or dst Is Nothing Then Exit Function
        
        Dim udraw As String
        
'        udraw  = 'l("%08x->%08x",'                  % (self.src, self.dst)
'        udraw +=   'e("",'                          # open edge
'        udraw +=     '['                            # open attributes
'        udraw +=       'a("EDGECOLOR","#%06x"),'    % self.color
'        udraw +=       'a("OBJECT","%s")'           % self.label
'        udraw +=     '],'                           # close attributes
'        udraw +=     'r("%08x")'                    % self.dst
'        udraw +=   ')'                              # close edge
'        udraw += ')'                                # close element

        'udraw = "l(""" & Hex(src.id) & "->" & Hex(dst.id) & """," & _
                 "e(""""," & _
                 "[" & _
                 "a(""EDGECOLOR"",""#" & Hex(color) & """)," & _
                 "a(""OBJECT"",""" & label & """)" & _
                 "]," & _
                 "r(""" & Hex(dst.id) & """)" & _
                 ")" & _
                 ")"
                 
        udraw = "l(""" & src.id & "->" & dst.id & """," & _
                 "e(""""," & _
                 "[" & _
                 "a(""OBJECT"",""" & label & """)" & _
                 "]," & _
                 "r(""" & dst.id & """)" & _
                 ")" & _
                 ")"
                 
        render_edge_udraw = udraw

End Function

'    ####################################################################################################################
'    def render_edge_udraw_update(self):
'        '''
'        Render an edge update description suitable for use in a GML file using the set internal attributes.
'
'        @rtype:  String
'        @return: GML edge update description
'        '''
'
'        udraw  = 'new_edge("%08x->%08x","",'      % (self.src, self.dst)
'        udraw +=   '['
'        udraw +=     'a("EDGECOLOR","#%06x"),'    % self.color
'        udraw +=       'a("OBJECT","%s")'         % self.label
'        udraw +=   '],'
'        udraw +=   '"%08x","%08x"'                % (self.src, self.dst)
'        udraw += ')'
'
'        return udraw
