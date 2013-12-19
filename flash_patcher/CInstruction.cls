VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Offset As Long
Public OpCodes As String
Public Disasm As String
Public InstructionLength As Long
Public RawOffset As Long
Public isLabel As Boolean
Public isPossibleBranch As Boolean
Public branchTarget As String
Public variableName As String

Public Label  As String
'

Function LoadLine(x, block As CAbcBlock) As Boolean
    
    On Error GoTo hell
    
    x = Trim(x)
    x = Replace(x, vbTab, " ")
    
    a = InStr(x, ":")
    If a > 0 Then
        b = Mid(x, 1, a - 1)
        b = Trim(Replace(b, "offset", ""))
        RawOffset = CLng(b)
        Offset = RawOffset + block.CodeOffset
    End If
    
    x = Mid(x, a + 1)
    a = InStr(x, "  ")
    OpCodes = Trim(Mid(x, 1, a))
    Disasm = Trim(Mid(x, a))
    Disasm = Replace(Disasm, "  ", " ")
    Disasm = Replace(Disasm, Chr(10), Empty)
    
    tmp = Split(OpCodes, " ")
    InstructionLength = UBound(tmp) + 1
    
    If VBA.Left(Disasm, 1) = "L" Then
        isLabel = True
        Label = Mid(Disasm, 1, InStr(Disasm, ":") - 1)
    End If
    
    'If InStr(Disasm, "L69") > 0 Then Stop
    a = InStrRev(Disasm, "L")
    If a > 0 Then
        tmp = Trim(Mid(Disasm, a + 1))
        If isDecimalNumber(tmp) Then
            isPossibleBranch = True
            branchTarget = "L" & tmp
        End If
    End If
        
    If InStr(Disasm, "builtin:") > 0 Then
        a = InStr(Disasm, "http:")
        If a > 0 Then
            URL = Mid(Disasm, a, InStrRev(Disasm, ":") - a + 1)
            Disasm = Replace(Disasm, URL, Empty)
        End If
    End If
    
    If InStr(Disasm, ":EventDispatcher:") > 0 Then
        a = InStr(Disasm, " ")
        If a > 0 Then
            URL = Mid(Disasm, a, InStrRev(Disasm, ":") - a + 1)
            Disasm = Replace(Disasm, URL, Empty)
        End If
    End If
    
    If InStr(Disasm, "getlocal") > 0 Or InStr(Disasm, "setlocal") > 0 Then
        a = InStrRev(Disasm, " ")
        If a > 0 Then
            Disasm = Mid(Disasm, 1, a) & "var_" & Mid(Disasm, a + 1)
            variableName = Mid(Disasm, a + 1)
        Else
            a = InStr(Disasm, "local") 'get/setlocal0-4 type
            If a > 0 Then
                a = Mid(Disasm, a + 5)
                If IsNumeric(a) Then
                    variableName = "var_" & a
                    Disasm = Replace(Disasm, "local" & a, " " & variableName)
                End If
            End If
        End If
    End If
    
    LoadLine = True
    DoEvents
Exit Function

hell:
    LoadLine = False
    Debug.Print "Error in CInst.Loadline(" & x
    
End Function


