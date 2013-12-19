VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'offset 2606: 01 02 04 09 0A 25    function :July:::UintToDouble(:uint, :uint)::Number
Public Prototype As String
Public StartOffset As Long
Public RawOffset As Long
Public RawBody As String
Public Instructions As New Collection
Public CodeLength As Long
Public Labels As New Collection
Public startBytes As String
Public varTypes As New Collection 'of CVariable
Public globals As New Collection


Public Function FindLabel(Label As String) As CInstruction
    
    Dim ci As CInstruction
    For Each ci In Labels
        If ci.Label = Label Then
            Set FindLabel = ci
            Exit Function
        End If
    Next
    
End Function


Function LoadSelf(proto, body, owner As CAbcBlock)
    
    On Error GoTo hell
    Dim x As String
    
    proto = Trim(proto)
    proto = Replace(Replace(proto, vbCr, ""), vbLf, "")
    
    a = InStr(proto, ":")
    If a > 0 Then
        b = InStr(a, proto, "function")
        startBytes = Trim(Mid(proto, a + 1, b - a - 1))
    End If
    
    If a > 0 Then
        b = Mid(proto, 1, a - 1)
        b = Trim(Replace(b, "offset", ""))
        RawOffset = CLng(b)
        
        If owner.SHIFT_CONST_K = 0 Then 'first time (hopefull not 0 always!)
            owner.SetKConstant RawOffset, startBytes
        End If

        StartOffset = RawOffset + owner.CodeOffset '+ owner.SHIFT_CONST_K
    End If
    
    a = InStr(proto, "function")
    If a > 0 Then
        b = Mid(proto, a)
        Prototype = Trim(b)
    End If
    
    Dim ci As CInstruction
    Dim vt As CVariable
    Dim lastCoerce As String
    
    tmp = Split(body, vbCrLf)
    For i = 2 To UBound(tmp)
        If InStr(tmp(i), "Entries") < 1 And InStr(tmp(i), "Extras") < 1 And InStr(tmp(i), "Symbol") < 1 Then
            Set ci = New CInstruction
            If ci.LoadLine(tmp(i), owner) Then
                Instructions.Add ci
                If ci.isLabel Then Labels.Add ci
                If Len(ci.variableName) > 0 And Len(lastCoerce) > 0 Then
                    If objKeyExistsInCollection(varTypes, ci.variableName) = False Then
                        Set vt = New CVariable
                        vt.varName = ci.variableName
                        vt.varType = lastCoerce
                        varTypes.Add vt, ci.variableName
                    End If
                    lastCoerce = Empty
                End If
            
                a = InStr(ci.Disasm, "coerce")
                If a > 0 Then
                     lastCoerce = Trim(Mid(ci.Disasm, a + 6))
                End If
                
                a = InStr(ci.Disasm, "convert")
                If a > 0 Then
                     lastCoerce = Trim(Mid(ci.Disasm, a + 7))
                End If
                
                Select Case lastCoerce
                    Case "_u": lastCoerce = "uint"
                    Case "_i": lastCoerce = "int"
                    Case "_d": lastCoerce = "double"
                    Case "_b": lastCoerce = "boolean"
                    Case "_s": lastCoerce = "String"
                    Case "_a": lastCoerce = "Variant"
                End Select
                
                a = InStr(ci.Disasm, "getproperty")
                If a < 1 Then a = InStr(ci.Disasm, "initproperty")
                
                If a > 0 Then
                    x = Mid(ci.Disasm, InStr(ci.Disasm, ":") + 1)
                    ci.variableName = x
                    If objKeyExistsInCollection(globals, x) = False Then
                        Set vt = New CVariable
                        vt.varName = x
                        globals.Add vt, x
                    End If
                End If
            
            End If
        End If
    Next
    
    'using last CInstruction offset and size to determine length of function..
    CodeLength = ci.Offset + ci.InstructionLength - StartOffset
    If CodeLength < 1 Then Stop
    
    Exit Function
hell:
    Debug.Print "Error in CFunc.LoadSelf: " & proto & " " & Err.Description
    
End Function

