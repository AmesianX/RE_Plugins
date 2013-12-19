VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Functions As Collection 'of CFunction
Dim fdata As String
Public dbg As ListBox

Public MainSWF As String
Private LastModified As disasmTimes

Public CurrentDecompilation As String
Public CurSWF As String
Public DecompressedSWF As String
Public StringPool As New Collection
Public blocks As New Collection 'of abcblock
Public BinaryData As New Collection


Private Type disasmTimes
    main As Date
    decomp As Date
    Disasm As Date
End Type

Private Sub LoadCfg(fpath As String)
    
    Dim d As String, f As String
    
    d = GetParentFolder(fpath)
    f = d & "\" & GetBaseName(fpath) & ".cfg"
    
    If FileExists(f) Then
        ff = FreeFile
        Open f For Binary As ff
        Get ff, , LastModified
        Close ff
    Else
        With LastModified
            .decomp = 0
            .Disasm = 0
            .main = 0
        End With
    End If
    
End Sub
    
Private Sub SaveCfg(fpath As String, decomp As String, Disasm As String)
    
    Dim d As String, f As String
    d = GetParentFolder(fpath)
    f = d & "\" & GetBaseName(fpath) & ".cfg"
    
    If FileExists(f) Then Kill f
        
    If FileExists(fpath) Then LastModified.main = FileDateTime(fpath)
    If FileExists(decomp) Then LastModified.decomp = FileDateTime(decomp)
    If FileExists(Disasm) Then LastModified.Disasm = FileDateTime(Disasm)
    
    ff = FreeFile
    Open f For Binary As ff
    Put ff, , LastModified
    Close ff
     
    
End Sub


Function Decompile(fpath As String) As Boolean
    
    Dim decomp As String
    Dim outFile As String
    Dim exe As String
    Dim swfDump As String
    Dim output As String
    
    kWarned = False
    Set Functions = New Collection
    Set StringPool = New Collection
    Set blocks = New Collection
    Set BinaryData = New Collection
    
    If MainSWF <> fpath Then LoadCfg fpath
    
    MainSWF = fpath
    CurrentDecompilation = Empty
    DecompressedSWF = Empty
    CurSWF = Empty
    dbg.Clear
    
    exe = Form1.txtFlexBin & "\swfdump.exe"
    
    If Not FolderExists(Form1.txtFlexBin) Then
        MsgBox "Flex folder not found in options"
        Exit Function
    End If
    
    If Not FileExists(exe) Then
        MsgBox "swfdump.exe not found install flex sdk"
        Exit Function
    End If
    
    If Not FileExists(fpath) Then
        MsgBox "File not found"
        Exit Function
    End If
    
    CurSWF = fpath
    fpath = GetShortName(fpath)
    outFile = fpath & ".disasm.txt"
    exe = GetShortName(exe)
    DecompressedSWF = Decompress(fpath)
    
    If LastModified.main = 0 Then
        LastModified.main = FileDateTime(fpath)
    Else
        Dim cur As Date
        cur = FileDateTime(fpath)
        If cur > LastModified.main Then
            dbg.AddItem "Parent SWF has been modified since last disasm, refreshing..."
            If FileExists(outFile) Then Kill outFile
            If FileExists(DecompressedSWF) Then Kill DecompressedSWF
            DecompressedSWF = Decompress(fpath)
            LastModified.main = cur
        Else
            cur = FileDateTime(DecompressedSWF)
            If cur > LastModified.decomp Then
                dbg.AddItem "Decompressed SWF has been modified since last disasm, refreshing..."
                If FileExists(outFile) Then Kill outFile
                LastModified.decomp = cur
            End If
        End If
    End If
    
    dbg.AddItem "Opening " & fpath
    
    If FileExists(outFile) Then
        Decompile = True
        CurrentDecompilation = outFile
        dbg.AddItem "Found existing decompilation using " & outFile
        Exit Function
    End If
    
    If Len(DecompressedSWF) = 0 Then
        MsgBox "Failed to get decompressed version see debug log", vbInformation
        Exit Function
    End If
    
    'todo check file exits and getshortname
    'exe = GetShortName("D:\_Lilguys\flash tools\flex_sdk_4.6\bin\swfdump.exe") 'output appears safe to use same as 4.1 version..
    swfDump = exe & " -showbytecode -showoffset -abc -out " & outFile & " " & DecompressedSWF
    output = CmdOutput.GetCommandOutput(swfDump, , True)
    
    If FileExists(outFile) Then
        If FileLen(outFile) = 0 Then
            dbg.AddItem output
            Exit Function
        End If
        Decompile = True
        CurrentDecompilation = outFile
        dbg.AddItem "new decompilation successful: " & outFile
    End If
        
    SaveCfg fpath, DecompressedSWF, outFile
    
End Function

Function LoadDecompilation() As Boolean

    Set Functions = New Collection
    Set StringPool = New Collection
    Set blocks = New Collection
    
    Dim f As New CFunction
    Dim spi As CStringPoolItem
    
    If Not FileExists(CurrentDecompilation) Then Exit Function
    
    fdata = ReadFile(CurrentDecompilation)
    
    Dim abc As CAbcBlock
    Dim bd As CBinaryData
    
    a = InStr(fdata, "<ProductInfo")
    If a > 0 Then
        b = InStr(a, fdata, vbCrLf)
        If b > 0 Then
            a = a + 1
            dbg.AddItem Replace(Mid(fdata, a, b - a - 2), "'", Empty)
        End If
    End If
    
    a = 1
    While a
        a = InStr(a, fdata, "<DefineBinaryData")
        If a > 0 Then
            Set bd = New CBinaryData
            c = InStrRev(fdata, "offset=", a) + 7
            b = InStr(c, fdata, " ")
            bd.Offset = CLng(Trim(Mid(fdata, c, b - c))) + 12 'sizeof(DefineBinaryDataTag)
            c = InStr(a, fdata, "id=") + 3
            b = InStr(c, fdata, " ")
            bd.ID = CLng(Trim(Replace(Mid(fdata, c, b - c), "'", Empty)))
            c = InStr(a, fdata, "length=") + 7
            b = InStr(c, fdata, " ")
            bd.Size = CLng(Trim(Replace(Mid(fdata, c, b - c), "'", Empty)))
            dbg.AddItem "Embedded Binary data found ID:" & bd.ID & " Offset: 0x" & Hex(bd.Offset) & " Size: 0x" & Hex(bd.Size)
            BinaryData.Add bd
            a = a + 1
        End If
    Wend
    
    a = 1
    i = 1
    While a
        a = InStr(a + 1, fdata, "<DoABC2>")
        If a > 0 Then
            b = InStr(a, fdata, "</DoABC2>")
            If b > 0 Then
                Set abc = New CAbcBlock
                abc.Index = i
                abc.StartAt = a
                abc.EndAt = b
                abc.Data = Mid(fdata, a, b - a)
                c = InStrRev(fdata, "offset=", a) + 7
                b = InStr(c, fdata, " ")
                abc.CodeOffset = CLng(Trim(Mid(fdata, c, b - c)))
                a = a + b + 1
                dbg.AddItem "DoABC block found at " & abc.StartAt & " CodeOffset: " & abc.CodeOffset
                abc.LoadSelf
                blocks.Add abc
                i = i + 1
            End If
        End If
    Wend
           
    LoadDecompilation = True
     
End Function




Private Function Decompress(fpath As String) As String
    
    Dim f As Long
    Dim b(3) As Byte
    
    
    Dim output As String
    Dim defDecomp As String
    
    defDecomp = fpath & ".decomp"
    
    If Not FileExists(fpath) Then Exit Function
    
    If FileExists(defDecomp) Then 'already decompressed one found use that...
        Decompress = defDecomp
        Exit Function
    End If
    
    f = FreeFile
    Open fpath For Binary As f
    Get f, , b()
    Close f
    
    If b(0) = Asc("C") And b(1) = Asc("W") And b(2) = Asc("S") Then
        If UnCompressCWS(fpath, defDecomp) Then
            Decompress = defDecomp
        Else
            dbg.AddItem "UnCompressCWS failed"
        End If
    Else
        FileCopy fpath, defDecomp
        dbg.AddItem "File not compressed continuing..."
        Decompress = defDecomp
    End If
    
    
End Function


Function UnCompressCWS(fpath, outFile) As Boolean
    
    On Error Resume Next
    Dim ff As Long
    Dim Header() As Byte
    Dim b() As Byte
    Dim bOut() As Byte
    
    ff = FreeFile
    ReDim Header(0 To 7)
    
    Open fpath For Binary As ff
    Get ff, , Header()
    
    If Header(0) <> Asc("C") Or Header(1) <> Asc("W") Or Header(2) <> Asc("S") Then
        dbg.AddItem "Can not decompress file does not have the CWS header"
        Exit Function
    End If
    
    ReDim b(LOF(ff) - 9)
    Get ff, , b()
    Close ff
    
    If Not modZlib.UncompressData(b(), bOut()) Then
        dbg.AddItem "Decompression Failed"
        Exit Function
    End If
    
    Header(0) = Asc("F")
    ff = FreeFile
    
    If FileExists(CStr(outFile)) Then Kill outFile
    
    Open outFile For Binary As ff
    Put ff, , Header()
    Put ff, , bOut()
    Close ff
    
    UnCompressCWS = True
    
End Function







