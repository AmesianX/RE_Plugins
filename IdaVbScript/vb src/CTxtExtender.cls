VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTxtExtender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'dzzie@yahoo.com
'http://sandsprite.com

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageStr Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetWindowDC Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "User32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LockWindowUpdate Lib "User32" (ByVal hwndLock As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCaret Lib "User32" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function ShowCaret Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "User32" () As Long
Private Declare Function SetCaretBlinkTime Lib "User32" (ByVal wMSeconds As Long) As Long
Private Declare Function GetCaretBlinkTime Lib "User32" () As Long

Private Type Rect
    left As Long
    Top As Long
    right As Long
    Bottom As Long
End Type

Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Public Enum tmMsgs
        EM_UNDO = &HC7
        EM_CANUNDO = &HC6
        EM_SETWORDBREAKPROC = &HD0
        EM_SETTABSTOPS = &HCB
        EM_SETSEL = &HB1
        EM_SETRECTNP = &HB4
        EM_SETRECT = &HB3
        EM_SETREADONLY = &HCF
        EM_SETPASSWORDCHAR = &HCC
        EM_SETMODIFY = &HB9
        EM_SCROLLCARET = &HB7
        EM_SETHANDLE = &HBC
        EM_SCROLL = &HB5
        EM_REPLACESEL = &HC2
        EM_LINESCROLL = &HB6
        EM_LINELENGTH = &HC1
        EM_LINEINDEX = &HBB
        EM_LINEFROMCHAR = &HC9
        EM_LIMITTEXT = &HC5
        EM_GETWORDBREAKPROC = &HD1
        EM_GETTHUMB = &HBE
        EM_GETRECT = &HB2
        EM_GETSEL = &HB0
        EM_GETPASSWORDCHAR = &HD2
        EM_GETMODIFY = &HB8
        EM_GETLINECOUNT = &HBA
        EM_GETLINE = &HC4
        EM_GETHANDLE = &HBD
        EM_GETFIRSTVISIBLELINE = &HCE
        EM_FMTLINES = &HC8
        EM_EMPTYUNDOBUFFER = &HCD
        EM_SETMARGINS = &HD3
End Enum

Private Const WM_VScroll = &H115
Private Const WM_CHAR = &H102
Private Const EC_LEFTMARGIN = &H1
Private Const EC_RIGHTMARGIN = &H2

Private myTopLine As Long
Private TrackingScroll As Boolean
Private OverRidingTabs As Boolean
Private OverrideTabNow As Boolean

Public WithEvents mTextBox As TextBox
Attribute mTextBox.VB_VarHelpID = -1

Public TabWidth As Integer
Public AutoIndent As Boolean

Public AddAccelerators As Boolean

Event Scrolled()
Event AutoComplete()
Event NewLine()



Function CurrentColumn() As Long
       CurrentColumn = mTextBox.selstart - SendMessage(mTextBox.hWnd, EM_LINEINDEX, -1&, ByVal 0&) + 1
End Function

Function CurrrentLineIndex() As Long
    CurrrentLineIndex = SendMessage(mTextBox.hWnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1
End Function

Function GetLine(lineNo As Long) As String
    'When retrieving a line into a buffer the first word
    'of the buffer specifies the max number of bytes to read,
    'so one has to guess the maximum line length.  When the bytes
    'are read into the buffer the first word gets overwritten.
    'Remember that lines in a TextBox are numbered starting at zero.

    Dim lret As Long
    Dim strBuffer As String
    Dim intLineLen As Integer
    Const MAX_LINE_LEN = 300
    
    strBuffer = Space(MAX_LINE_LEN) 'max width possible for line
    intLineLen = SendMessageStr(mTextBox.hWnd, EM_GETLINE, lineNo, strBuffer)
    GetLine = left(strBuffer, intLineLen)


End Function


Property Get LineCount() As Long
    LineCount = SendMessage(mTextBox.hWnd, EM_GETLINECOUNT, 0, ByVal 0&)
End Property

Property Get TopLineIndex() As Long
    TopLineIndex = SendMessage(mTextBox.hWnd, EM_GETFIRSTVISIBLELINE, 0, ByVal 0&)
End Property

Property Get VisibleLines() As Long
    Dim udtRect As Rect, tm As TEXTMETRIC
    Dim hdc As Long, lFont As Long, lOrgFont As Long
    Const WM_GETFONT As Long = &H31
    
    SendMessage mTextBox.hWnd, EM_GETRECT, 0, udtRect

    lFont = SendMessage(mTextBox.hWnd, WM_GETFONT, 0, 0)
    hdc = GetDC(mTextBox.hWnd)

    If lFont <> 0 Then
        lOrgFont = SelectObject(hdc, lFont)
    End If

    GetTextMetrics hdc, tm
    
    If lFont <> 0 Then
        lFont = SelectObject(hdc, lOrgFont)
    End If

    VisibleLines = (udtRect.Bottom - udtRect.Top) \ tm.tmHeight

    ReleaseDC mTextBox.hWnd, hdc

End Property

Sub ScrollToLine(x As Integer)
     x = x - TopLineIndex
     ScrollIncremental , x
End Sub

Sub ScrollIncremental(Optional horz As Integer = 0, Optional vert As Integer = 0)
    'lParam&  The low-order 2 bytes specify the number of vertical
    '          lines to scroll. The high-order 2 bytes specify the
    '          number of horizontal columns to scroll. A positive
    '          value for lParam& causes text to scroll upward or to the
    '          left. A negative value causes text to scroll downward or
    '          to the right.
    ' r&       Indicates the number of lines actually scrolled.
    
    Dim r As Long
    r = CLng(&H10000 * horz) + vert
    r = SendMessage(mTextBox.hWnd, EM_LINESCROLL, 0, ByVal r)

End Sub

Private Sub mTextBox_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim tli As Long
    Dim lstLine As String, spacer As String
    
    tli = TopLineIndex
    If tli <> myTopLine Then  'textbox has auto scrolled on us
        myTopLine = tli
        RaiseEvent Scrolled
    End If
   
    If KeyCode = 13 And AutoIndent Then
        LockWindowUpdate mTextBox.hWnd
        tli = CurrrentLineIndex
        lstLine = GetLine(tli - 2)
        While left(lstLine, 1) = " " Or left(lstLine, 1) = vbTab
            spacer = spacer & left(lstLine, 1)
            lstLine = Mid(lstLine, 2, Len(lstLine))
        Wend
        mTextBox.SelText = spacer
        LockWindowUpdate 0
    End If
        
    'Form1.Caption = KeyCode & " " & Shift
    
    If AddAccelerators Then
    
        If KeyCode = 65 And Shift = 2 Then 'select all ctrl-a
            mTextBox.selstart = 0
            mTextBox.SelLength = Len(mTextBox.Text)
        End If
        
        If KeyCode = 32 And Shift = 2 Then 'autocomplete ctrl-space
            RaiseEvent AutoComplete
        End If
        
    End If
    
    If KeyCode = 13 Then RaiseEvent NewLine
    
End Sub


