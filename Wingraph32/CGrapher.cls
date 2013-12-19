VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGrapher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public nodes As New Collection
Public NodeWidth As Long
Public NodeHeight As Long
Public afControl As AddFlow
Public layout As HFlow
Public Colors As Collection
Public fxLv As ListView

Dim fso As New CFileSystem2

Function LoadFile(pth As String)
    
    On Error Resume Next
    
    If Not fso.FileExists(pth) Then Exit Function
    
    Set Colors = New Collection
    
    Dim n As afNode
    Dim sNode As afNode
    Dim tNode As afNode
    Dim link As afLink
    Dim li As ListItem
    
    If NodeWidth = 0 Then NodeWidth = 1000
    If NodeHeight = 0 Then NodeHeight = 1000
    
    tmp = fso.ReadFile(pth)
    tmp = Split(tmp, vbCrLf)
    
    For Each x In tmp
        If VBA.Left(x, Len("node:")) = "node:" Then
            'add a node  title: "0" label: "sub_4122CC" color: 76 textcolor: black
            t = GetParam(x, "title")
            l = GetParam(x, "label")
            c = Get_NQ_Param(x, "color")
            tc = Get_NQ_Param(x, "textcolor")
            
            If Len(t) > 0 And Len(l) > 0 Then
                Set n = afControl.nodes.Add(0, 0, NodeWidth, NodeHeight)
                Set li = fxLv.ListItems.Add(, , l)
                Set li.Tag = n
                'If InStr(l, "memcpy") > 0 Then Stop
                n.Text = Space(3) & l & Space(3)
                n.Tag = t
                n.ForeColor = getColorFromId(tc)
                If c <> vbWhite Then n.FillColor = getColorFromId(c)
                If n.ForeColor = n.FillColor Then n.ForeColor = vbWhite
                nodes.Add n, "node:" & t
            End If
            
        ElseIf VBA.Left(x, Len("edge:")) = "edge:" Then
            'add a link
            s = GetParam(x, "sourcename")
            t = GetParam(x, "targetname")
            If Len(s) > 0 Then Set sNode = GetNodeID(s)
            If Len(t) > 0 Then Set tNode = GetNodeID(t)
            If Not sNode Is Nothing And Not tNode Is Nothing Then
                Set link = sNode.Links.Add(tNode)
                link.Selectable = False
            End If
        ElseIf VBA.Left(x, Len("colorentry")) = "colorentry" Then
            'colorentry 32: 0 0 0
            cnum = Mid(x, Len("colorentry") + 2, 2)
            cset = Trim(Mid(x, InStr(x, ":") + 1))
            Colors.Add toColor(cset), "cid:" & cnum
        End If
    Next
    
    'layout.xMargin = 1000
    'layout.yMargin = 1000
    'layout.VertexDistance = 10000
    layout.layout afControl
    
    If Err.Number <> 0 Then
        MsgBox "Error in CGrapher.LoadFile: " & Err.Description, vbInformation
    End If
    
    Exit Function

End Function

Private Function getColorFromId(id) As Long
    On Error Resume Next
    
    getColorFromId = Colors("cid:" & id)
    If Err.Number <> 0 Then
        getColorFromId = id
    End If
    
    If getColorFromId = vbYellow Then getColorFromId = vbBlue 'fuckyou

End Function
Private Function toColor(s) As Long
        
    On Error Resume Next
    t = Split(s, " ")
    For i = 0 To UBound(t)
        t(i) = Hex(t(i))
        If Len(t(i)) = 1 Then t(i) = "0" & t(i)
    Next
    
    toColor = CLng("&h" & t(0) & t(1) & t(2))
   ' If toColor <> 0 Then Stop
    

End Function

Function GetNodeID(id) As afNode
    On Error Resume Next
    Set GetNodeID = nodes("node:" & id)
End Function

Public Function Get_NQ_Param(src, param)
    'only works on NON-quoted values works for our needs
    'node: { title: "0" label: "sub_4122CC" color: 76 textcolor: 73 borderwidth: 10 bordercolor: 82  }
    'edge: { sourcename: "1" targetname: "0" }
    
    a = InStr(1, src, param, vbTextCompare)
    If a < 1 Then Exit Function 'parameter not found
    
    a = a + 2
    b = InStr(a, src, " ") + 1
    If b < 1 Then Exit Function 'next space not found.
    
    c = InStr(b, src, " ")
    If c < 1 Then Exit Function
    If c < b Then Exit Function
    
    Get_NQ_Param = Mid(src, b, c - b)
    
    If Get_NQ_Param = "white" Then Get_NQ_Param = vbWhite
    If Get_NQ_Param = "black" Then Get_NQ_Param = vbBlack
    
End Function


Public Function GetParam(src, param)
    'only works on quoted values works for our needs
    'node: { title: "0" label: "sub_4122CC" color: 76 textcolor: 73 borderwidth: 10 bordercolor: 82  }
    'edge: { sourcename: "1" targetname: "0" }
    
    a = InStr(1, src, param, vbTextCompare)
    If a < 1 Then Exit Function 'parameter not found
    
    b = InStr(a, src, """") + 1
    If b < 1 Then Exit Function 'quote not found.
    
    c = InStr(b, src, """")
    If c < 1 Then Exit Function
    If c < b Then Exit Function
    
    GetParam = Mid(src, b, c - b)
    
End Function
