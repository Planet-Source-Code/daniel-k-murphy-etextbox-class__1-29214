VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsETextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ------------------------------------------------------------
' eTextBox.cls
'
' Programmed By: Daniel K Murphy
'
' Started On: November 13, 2001
' Completed On: November 25, 2001
' Updated On: November 26, 2001
'
' Portions of code written by other Authors --
'    Noted in the code snippets where used
'
' ... My "Thanks" to them!
' ------------------------------------------------------------

' ----------------------------------
' Purpose:
'   To allow the use of the simple
' TextBox Control to have the
' abilities Below:
'
' Count total numbers of characters,
' Which line the cursor is on,
' How many lines, Top Visible Line,
' Bottom Visible line, How many
' lines the Text Box is able to show,
' Total Word Count, LineIndex,
' LineLength, LineData, Undo, CanUndo,
' ClearUndoBuffer, SoftBreaks,
' IsDirty, UnSelect, CaretPos,
' and ReadOnly
' ----------------------------------

' ----------------------------------
' Original intent was to give the
' capabilities of the Rich Text Box,
' but due to the lack of ability to
' Add multiple colors to a TextBox,
' and to have multiple fonts, I was
' unable to add these features.
'
' If anyone is aware of how this can
' be done using a TextBox, please
' drop me an email.
'
' daniel_k_murphy@hotmail.com
'
' ----------------------------------

' ----------------------------------
' Declare the API Functions
' ----------------------------------
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageBynum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageBystring Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

' ----------------------------------
' Declare the API Subroutines
' ----------------------------------
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub PokeLng Lib "kernel32" Alias "RtlMoveMemory" (ByVal Addr As Long, Value As Long, Optional ByVal nBytes As Long = 4)

' ----------------------------------
' Declare the TYPES
' ----------------------------------
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

' ----------------------------------
' This is used for the INFO on the TextBox
' ----------------------------------
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

' ----------------------------------
' Declare the Constants
' ----------------------------------

' Edit Control Undo Messages
Private Const EM_CANUNDO = &HC6
Private Const EM_UNDO = &HC7
Private Const EM_EMPTYUNDOBUFFER = &HCD

' Edit Control Text Formatting Messages
Private Const EM_FMTLINES = &HC8
Private Const EM_GETRECT = &HB2
Private Const EM_LIMITTEXT = &HC5
Private Const EM_SETRECT = &HB3
Private Const EM_SETRECTNP = &HB4
Private Const EM_SETTABSTOPS = &HCB

' Edit Control Selection and Display Messages
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_GETLINE = &HC4
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETMODIFY = &HB8
Private Const EM_GETPASSWORDCHAR = &HD2
Private Const EM_GETSEL = &HB0
Private Const EM_GETWORDBREAKPROC = &HD1
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_LINESCROLL = &HB6
Private Const EM_REPLACESEL = &HC2
Private Const EM_SETMODIFY = &HB9
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const EM_SETREADONLY = &HCF
Private Const EM_SETSEL = &HB1
Private Const EM_SETWORDBREAKPROC = &HD0

' Edit Control Scroll Messages
Private Const EM_GETTHUMB = &HBE
Private Const EM_SCROLL = &HB5
Private Const EM_SCROLLCARET = &HB7

' Edit Control Window Messages
Private Const EM_GETHANDLE = &HBD
Private Const EM_SETHANDLE = &HBC

' General Windows(OS) Messages
' This doesn't need them ALL, but they are all here
Private Const WM_NULL = &H0
Private Const WM_CREATE = &H1
Private Const WM_DESTROY = &H2
Private Const WM_MOVE = &H3
Private Const WM_SIZE = &H5
Private Const WM_ACTIVATE = &H6
Private Const WM_SETFOCUS = &H7
Private Const WM_KILLFOCUS = &H8
Private Const WM_ENABLE = &HA
Private Const WM_SETREDRAW = &HB
Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_PAINT = &HF
Private Const WM_CLOSE = &H10
Private Const WM_QUERYENDSESSION = &H11
Private Const WM_QUIT = &H12
Private Const WM_QUERYOPEN = &H13
Private Const WM_ERASEBKGND = &H14
Private Const WM_SYSCOLORCHANGE = &H15
Private Const WM_ENDSESSION = &H16
Private Const WM_SYSTEMERROR = &H17
Private Const WM_SHOWWINDOW = &H18
Private Const WM_CTLCOLOR = &H19
Private Const WM_WININICHANGE = &H1A
Private Const WM_DEVMODECHANGE = &H1B
Private Const WM_ACTIVATEAPP = &H1C
Private Const WM_FONTCHANGE = &H1D
Private Const WM_TIMECHANGE = &H1E
Private Const WM_CANCELMODE = &H1F
Private Const WM_SETCURSOR = &H20
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_CHILDACTIVATE = &H22
Private Const WM_QUEUESYNC = &H23
Private Const WM_GETMINMAXINFO = &H24
Private Const WM_PAINTICON = &H26
Private Const WM_ICONERASEBKGND = &H27
Private Const WM_NEXTDLGCTL = &H28
Private Const WM_SPOOLERSTATUS = &H2A
Private Const WM_DRAWITEM = &H2B
Private Const WM_MEASUREITEM = &H2C
Private Const WM_DELETEITEM = &H2D
Private Const WM_VKEYTOITEM = &H2E
Private Const WM_CHARTOITEM = &H2F
Private Const WM_SETFONT = &H30
Private Const WM_GETFONT = &H31
Private Const WM_COMMNOTIFY = &H44
Private Const WM_QUERYDRAGICON = &H37
Private Const WM_COMPAREITEM = &H39
Private Const WM_COMPACTING = &H41
Private Const WM_WINDOWPOSCHANGING = &H46
Private Const WM_WINDOWPOSCHANGED = &H47
Private Const WM_POWER = &H48
Private Const WM_NCCREATE = &H81
Private Const WM_NCDESTROY = &H82
Private Const WM_NCCALCSIZE = &H83
Private Const WM_NCHITTEST = &H84
Private Const WM_NCPAINT = &H85
Private Const WM_NCACTIVATE = &H86
Private Const WM_GETDLGCODE = &H87
Private Const WM_NCMOUSEMOVE = &HA0
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_NCLBUTTONUP = &HA2
Private Const WM_NCLBUTTONDBLCLK = &HA3
Private Const WM_NCRBUTTONDOWN = &HA4
Private Const WM_NCRBUTTONUP = &HA5
Private Const WM_NCRBUTTONDBLCLK = &HA6
Private Const WM_NCMBUTTONDOWN = &HA7
Private Const WM_NCMBUTTONUP = &HA8
Private Const WM_NCMBUTTONDBLCLK = &HA9
Private Const WM_KEYFIRST = &H100
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_CHAR = &H102
Private Const WM_DEADCHAR = &H103
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105
Private Const WM_SYSCHAR = &H106
Private Const WM_SYSDEADCHAR = &H107
Private Const WM_KEYLAST = &H108
Private Const WM_INITDIALOG = &H110
Private Const WM_COMMAND = &H111
Private Const WM_SYSCOMMAND = &H112
Private Const WM_TIMER = &H113
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const WM_INITMENU = &H116
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_MENUSELECT = &H11F
Private Const WM_MENUCHAR = &H120
Private Const WM_ENTERIDLE = &H121
Private Const WM_MOUSEFIRST = &H200
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MOUSELAST = &H209
Private Const WM_PARENTNOTIFY = &H210
Private Const WM_MDICREATE = &H220
Private Const WM_MDIDESTROY = &H221
Private Const WM_MDIACTIVATE = &H222
Private Const WM_MDIRESTORE = &H223
Private Const WM_MDINEXT = &H224
Private Const WM_MDIMAXIMIZE = &H225
Private Const WM_MDITILE = &H226
Private Const WM_MDICASCADE = &H227
Private Const WM_MDIICONARRANGE = &H228
Private Const WM_MDIGETACTIVE = &H229
Private Const WM_MDISETMENU = &H230
Private Const WM_DROPFILES = &H233
Private Const WM_CUT = &H300
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Private Const WM_CLEAR = &H303
Private Const WM_UNDO = &H304
Private Const WM_RENDERFORMAT = &H305
Private Const WM_RENDERALLFORMATS = &H306
Private Const WM_DESTROYCLIPBOARD = &H307
Private Const WM_DRAWCLIPBOARD = &H308
Private Const WM_PAINTCLIPBOARD = &H309
Private Const WM_VSCROLLCLIPBOARD = &H30A
Private Const WM_SIZECLIPBOARD = &H30B
Private Const WM_ASKCBFORMATNAME = &H30C
Private Const WM_CHANGECBCHAIN = &H30D
Private Const WM_HSCROLLCLIPBOARD = &H30E
Private Const WM_QUERYNEWPALETTE = &H30F
Private Const WM_PALETTEISCHANGING = &H310
Private Const WM_PALETTECHANGED = &H311

' -----------------------------
' Declare the Private Variables
' -----------------------------
Private txtTextBox As TextBox
Private dblLineNumber As Double
Private dblLineCount As Double
Private dblVisibleLines As Double
Private dblTopLine As Double
Private dblBottomLine As Double
Private dblCharacterCount As Double
Private dblWordCount As Double
Private dblLineLength As Double
Private blnCanUndo As Boolean
Private blnSoftBreaks As Boolean
Private blnIsDirty As Boolean
Private lngCaretPos As Long
Private blnReadOnly As Boolean

' -----------------------------
' Create the Properties
' -----------------------------

' ----------------------------------------
' Set Properties
' ----------------------------------------
Public Property Set eTextBox(eText As TextBox)
    Set txtTextBox = eText
End Property

' ----------------------------------------
' Read Properties
' ----------------------------------------
Public Property Get LineNumber() As Double
    If txtTextBox Is Nothing Then
        ' No textbox is set
        ' Return -1
        dblLineNumber = -1
    Else
        ' Get the Position
        dblLineNumber = GetLineNumber
    End If
    
    ' Return Value
    LineNumber = dblLineNumber
End Property

Public Property Get LineCount() As Double
    If txtTextBox Is Nothing Then
        ' No textbox is set
        ' Return -1
        dblLineCount = -1
    Else
        ' Get the Count
        dblLineCount = GetNumberOfLines
    End If
    
    ' Return Value
    LineCount = dblLineCount
End Property

Public Property Get VisibleLines() As Double
    If txtTextBox Is Nothing Then
        ' No textbox is set
        ' Return -1
        dblVisibleLines = -1
    Else
        ' Get the Number of Visible Lines
        dblVisibleLines = GetVisibleLines
    End If
    
    ' Return Value
    VisibleLines = dblVisibleLines
End Property

Public Property Get TopLine() As Double
    If txtTextBox Is Nothing Then
        ' No textbox is set
        ' Return -1
        dblTopLine = -1
    Else
        ' Get the Number of Visible Lines
        dblTopLine = GetTopLine
    End If
    
    ' Return Value
    TopLine = dblTopLine
End Property

Public Property Get BottomLine() As Double
    If txtTextBox Is Nothing Then
        ' No textbox is set
        ' Return -1
        dblBottomLine = -1
    Else
        ' Get the Number of Visible Lines
        dblBottomLine = GetBottomLine
    End If
    
    ' Return Value
    BottomLine = dblBottomLine
End Property

Public Property Get CharacterCount() As Double
    If txtTextBox Is Nothing Then
        ' No textbox is set
        ' Return -1
        dblCharacterCount = -1
    Else
        ' Get the Number of Visible Lines
        dblCharacterCount = GetCharacterCount
    End If
    
    ' Return Value
    CharacterCount = dblCharacterCount
End Property

Public Property Get WordCount() As Double
    If txtTextBox Is Nothing Then
        ' No textbox is set
        ' Return -1
        dblWordCount = -1
    Else
        ' Get the Number of Visible Lines
        dblWordCount = GetWordCount
    End If
    
    ' Return Value
    WordCount = dblWordCount
End Property

Public Property Get CanUndo() As Boolean
    If txtTextBox Is Nothing Then
        ' No textbox is set
        ' Return False
        blnCanUndo = False
    Else
        ' Get the Number of Visible Lines
        blnCanUndo = GetCanUndo
    End If
    
    ' Return Value
    CanUndo = blnCanUndo
End Property

Public Property Get IsDirty() As Boolean
    If txtTextBox Is Nothing Then
        ' No TextBox is set
        ' Return False
        IsDirty = False
    Else
        ' Return result
        GetIsDirty
        IsDirty = blnIsDirty
    End If
End Property

Public Property Get CaretPos() As Double
    If txtTextBox Is Nothing Then
        ' No TextBox is set
        ' Return -1
        CaretPos = -1
    Else
        ' Get Position
        lngCaretPos = GetCaretPosition
        
        ' Return result
        CaretPos = lngCaretPos
    End If
End Property

Public Property Get ReadOnly() As Boolean
    ' Get Property Value
    ReadOnly = blnReadOnly
End Property

' -----------------------------
' Write Properties
' -----------------------------

Public Property Let SoftBreaks(SoftBreak As Boolean)
    ' Set property
    blnSoftBreaks = SoftBreak
    
    ' Change Textbox property
    SetSoftBreaks
End Property

Public Property Let ReadOnly(bln As Boolean)
    ' Set Property
    blnReadOnly = bln
    
    ' Change property
    SetReadOnly
End Property

' -----------------------------
' Public Subroutines
' -----------------------------
Public Sub ClearUndoBuffer()
    Dim lngClearUndoBuffer
    
    If Not (txtTextBox Is Nothing) Then
        ' Clear the Undo Buffer
        lngClearUndoBuffer = SendMessageBynum(txtTextBox.hwnd, EM_EMPTYUNDOBUFFER, 0, 0)
    End If
End Sub

Public Sub UnSelect()
    Dim lngUnSelect As Long
    
    If Not (txtTextBox Is Nothing) Then
        ' Clear the TextBox
        lngUnSelect = SendMessageBynum(txtTextBox.hwnd, WM_CLEAR, 0, 0)
    End If
End Sub

Public Sub LoadTEXT(FileName As String)
    Dim intFileNum As Integer
    Dim strText As String
    Dim strTemp As String
    
    ' Check to see if a FileName was given
    If Trim(FileName) <> "" Then
        ' Get the File Number
        intFileNum = FreeFile
        
        ' Open the File
        Open FileName For Input As #intFileNum
        
        ' Loop through each line, appending to the textbox
        Do Until EOF(intFileNum)
            ' Keep Windows Breathing
            DoEvents
            
            ' Get the data
            Line Input #intFileNum, strTemp
            
            ' Append to strText
            If strText = "" Then
                ' Nothing in the string
                strText = strTemp
            Else
                ' Something exists, start on a new line
                strText = strText + Chr$(13) + Chr$(10) + strTemp
            End If
        Loop
        
        ' This looks redundant, but it smooths out how
        ' the text is added to the textbox by first
        ' appending to a string, then placing into
        ' the txtTextBox.Text
        
        ' Append to the TextBox
        If txtTextBox = "" Then
            ' Nothing in the TextBox
            txtTextBox.Text = strText
        Else
            ' Something exists, start on a new line
            txtTextBox.Text = txtTextBox.Text + Chr$(13) + Chr$(10) + strText
        End If
        
        ' Close file
        Close #intFileNum
    End If
End Sub

Public Sub SaveTEXT(FileName As String)
    Dim intFileNum As Integer
    Dim strText As String
    
    ' Check to see if a FileName was given
    If Trim(FileName) <> "" Then
        ' Get the File Number
        intFileNum = FreeFile
        
        ' Open the File
        Open FileName For Output As #intFileNum
        
        ' Get the Text from the TextBox
        strText = txtTextBox.Text
        
        ' Write to file
        Print #intFileNum, Trim(strText)
        
        ' Close file
        Close #intFileNum
    End If
End Sub

Public Sub Clear()
    ' This sets the textbox.text to null
    ' Ok - OK!  So it isn't original,
    ' but why not?
    txtTextBox.Text = ""
End Sub

' -----------------------------
' Public Functions
' -----------------------------
Public Function LineLength(LineNum As Long) As Double
    Dim dblLineLength As Double
    
    If txtTextBox Is Nothing Then
        ' No TextBox is set
        ' Return -1
        dblLineLength = -1
    Else
        ' Get the value
        dblLineLength = GetLineLength(LineNum)
    End If
    
    ' Return Value
    LineLength = dblLineLength
End Function

Public Function LineData(LineNum As Long) As String
    Dim strLine As String
    Dim lngLineIndex As Long
    Dim lngLineLength As Long
    
    If txtTextBox Is Nothing Then
        ' No TextBox is set
        ' Return ERROR
        strLine = "ERROR"
    Else
        ' Get the Line Index
        lngLineIndex = CLng(GetLineIndex(LineNum))
        
        ' Get the Line Length
        lngLineLength = CLng(GetLineLength(LineNum))
        
        ' Set Selection
        With txtTextBox
            .SelStart = lngLineIndex
            .SelLength = lngLineLength
            strLine = .SelText
        End With
    End If
    
    ' Return value
    LineData = strLine
End Function

Public Function LineIndex(LineNum As Long) As Double
    Dim dblLineIndex As Double
    
    If txtTextBox Is Nothing Then
        ' No TextBox is Set
        ' Return -1
        dblLineIndex = -1
    Else
        ' Get Line Index
        dblLineIndex = GetLineIndex(LineNum)
    End If
    
    ' Return value
    LineIndex = dblLineIndex
End Function

Public Function Undo() As Boolean
    Dim blnCanUndo As Boolean
    Dim blnUndo As Boolean
    
    ' This will check to see if an Undo can take place
    blnCanUndo = GetCanUndo
    
    If blnCanUndo = True Then
        ' Undo is able to take place
        blnUndo = SendMessageBynum(txtTextBox.hwnd, EM_UNDO, 0, 0)
    Else
        ' False
        blnUndo = False
    End If
    
    Undo = blnUndo
End Function

' -----------------------------
' Private Functions
' -----------------------------
Private Function GetLineNumber() As Long
    Dim lngSelectedText As Long
    Dim lngLineNumber As Long
    
    ' This will return the Selected Text
    lngSelectedText = SendMessageBynum&(txtTextBox.hwnd, EM_GETSEL, 0, 0&)
    
    ' This will return the Actual Line number
    lngLineNumber = SendMessageBynum(txtTextBox.hwnd, EM_LINEFROMCHAR, lngSelectedText, 0&)
    
    ' Return the Position
    GetLineNumber = lngLineNumber
End Function

Private Function GetNumberOfLines() As Long
    Dim lngNumberOfLines As Long
    
    ' Get the number of Lines in the TextBox
    lngNumberOfLines = SendMessageBynum(txtTextBox.hwnd, EM_GETLINECOUNT, 0, 0&)
    
    ' Return the Count
    GetNumberOfLines = lngNumberOfLines
End Function

Private Function GetVisibleLines() As Long
    Dim rectText As RECT
    Dim lnghDC As Long
    Dim lngFont As Long
    Dim lngOldFont As Long
    Dim lngDI As Long
    Dim lngLC As Long
    Dim typTM As TEXTMETRIC

    ' Get the formatting rectangle - this describes the
    ' rectangle in the control in which text is placed.
    lngLC = SendMessage(txtTextBox.hwnd, EM_GETRECT, 0, rectText)

    ' Get a handle to the logical font used by the control.
    ' The VB font properties are accurately reflected by
    ' this logical font.
    lngFont = SendMessageBynum(txtTextBox.hwnd, WM_GETFONT, 0, 0&)
    
    ' Get a device context to the text control.
    lnghDC = GetDC(txtTextBox.hwnd)

    ' Select in the logical font to obtain the exact font
    ' metrics.
    If lngFont <> 0 Then lngOldFont = SelectObject(lnghDC, lngFont)

    lngDI = GetTextMetrics(lnghDC, typTM)
    ' Select out the logical font
    If lngFont <> 0 Then lngFont = SelectObject(lnghDC, lngOldFont)

    ' The lines depends on the formatting rectangle and font height
    GetVisibleLines = (rectText.bottom - rectText.top) / typTM.tmHeight
 
    ' Release the device context when done.
    lngDI = ReleaseDC(txtTextBox.hwnd, lnghDC)
End Function

Private Function GetTopLine() As Long
    Dim lngFirstVisibleLine As Long
    
    ' Get the first visible Line
    lngFirstVisibleLine = SendMessageBynum(txtTextBox.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0&)
    
    ' Return Value
    GetTopLine = lngFirstVisibleLine
End Function

Private Function GetBottomLine() As Long
    Dim lngBottomLine As Long
    Dim lngTopLine As Long
    Dim lngVisibleLines As Long
    Dim lngNumberOfLines As Long
    
    ' Get the top line
    lngTopLine = GetTopLine
    
    ' Get the Number of Visible Lines
    lngVisibleLines = GetVisibleLines
    
    ' Get TOTAL Number of lines
    lngNumberOfLines = GetNumberOfLines
    
    ' Determine Last Visible Line
    If lngTopLine = lngNumberOfLines - 1 Then
        ' Top line and Last line are equal
        lngBottomLine = lngTopLine
    Else
        ' It is other than the top line
        If lngTopLine + lngVisibleLines > lngNumberOfLines Then
            ' It is between the Top Line and the Bottom of the textbox
            ' Return the Number Of Lines
            lngBottomLine = lngNumberOfLines - 1
        Else
            ' It will be the LAST visible line in the TextBox
            lngBottomLine = lngTopLine + lngVisibleLines - 1
        End If
    End If
    
    ' Return Value
    GetBottomLine = lngBottomLine
End Function

Private Function GetCharacterCount() As Long
    Dim intASCII() As Integer
    Dim Ptr As Long
    Dim i As Long
  
    ' ----------------------------------
    ' Original Code Created By:
    '   Christopher D Lucas
    '
    ' I enhanced the code to meet my needs
    ' ----------------------------------
    
    ' Move intASCII's pointer to coincide with
    ' sText's character array in memory
    ReDim intASCII(1 To 1)
    PokeLng VarPtr(Ptr), ByVal ArrPtr(intASCII)
    PokeLng Ptr + 16, &HFFFFFFF
    PokeLng Ptr + 12, StrPtr(txtTextBox.Text)
    
    ' Now loop through intASCII and count matches
    For i = 1 To Len(txtTextBox.Text)
        If intASCII(i) > 32 Then
            GetCharacterCount = GetCharacterCount + 1
        End If
    Next i
End Function

Private Function GetWordCount() As Long
    Dim dest() As Byte
    Dim i As Long
    Dim cChar33 As Byte: cChar33 = 33
    Dim cChar64 As Byte: cChar64 = 64
    Dim cChar91 As Byte: cChar91 = 91
    Dim cChar96 As Byte: cChar96 = 96
    Dim cChar123 As Byte: cChar123 = 123
    
    Dim strText As String
    
    ' ----------------------------------
    ' Original Code Created By:
    '   Christopher D Lucas
    '
    ' Found on www.Planet-Source-Code.com
    '
    ' I enhanced the code to meet my needs
    ' ----------------------------------
    
    ' Get the text of the TextBox
    strText = Trim(txtTextBox.Text)
    
    ' See if there is something in the TextBox
    If LenB(strText) Then
    
        ' Move the string's byte array into dest()
        ReDim dest(LenB(strText))
        CopyMemory dest(0), ByVal StrPtr(strText), LenB(strText) - 1
        
        ' Now loop through the array and count the words
        For i = 0 To UBound(dest) Step 2
            
            ' Make sure that there are Letters inside the string
            If dest(i) > cChar96 Then
                If dest(i) < cChar123 Then
                    
                    ' Upper Case Letters
                    Do Until dest(i) < cChar33
                       i = i + 2
                    Loop
                    
                    ' Add one to the Word Count
                    GetWordCount = GetWordCount + 1
                End If
            ElseIf dest(i) > cChar64 Then
                If dest(i) < cChar91 Then
                    
                    ' Lower Case Letters
                    Do Until dest(i) < cChar33
                       i = i + 2
                    Loop
                    
                    ' Add one to the Word Count
                    GetWordCount = GetWordCount + 1
                End If
            End If
        Next i
        
        ' Remove the Array from Memory
        Erase dest
    End If
End Function

Private Function GetLineLength(LineNum As Long) As Double
    Dim dblLineLength As Double
    Dim lngNumberOfLines As Long
    Dim dblLineIndex As Double
    
    ' Make sure the TextBox Is Not Nothing
    If txtTextBox Is Nothing Then
        ' No TextBox is set
        ' Return -1
        dblLineLength = -1
    Else
        ' Get the number of lines
        lngNumberOfLines = GetNumberOfLines
        
        ' See if the LineNum is out of bounds
        If LineNum > lngNumberOfLines Then
            dblLineLength = -2
        Else
            ' Get the Index of the line First
            dblLineIndex = GetLineIndex(LineNum)
            
            ' Get the Line Length
            dblLineLength = SendMessageBynum(txtTextBox.hwnd, EM_LINELENGTH, dblLineIndex, 0)
        End If
    End If
    
    ' Return Value
    GetLineLength = dblLineLength
End Function

Private Function GetLineIndex(LineNum) As Long
    Dim lngLineIndex As Long
    
    ' Get the value
    lngLineIndex = SendMessageBynum(txtTextBox.hwnd, EM_LINEINDEX, LineNum, 0)
    
    ' Return Value
    GetLineIndex = lngLineIndex
End Function

Private Function GetCanUndo() As Boolean
    Dim lngCanUndo As Long
    
    ' Get the value
    lngCanUndo = SendMessageBynum(txtTextBox.hwnd, EM_CANUNDO, 0, 0)
    
    ' Return Value
    If lngCanUndo = 0 Then
        ' False
        GetCanUndo = False
    Else
        ' True
        GetCanUndo = True
    End If
End Function

Private Function GetIsDirty() As Boolean
    Dim lngIsDirty As Long
    
    ' Get value
    lngIsDirty = SendMessageBynum(txtTextBox.hwnd, EM_GETMODIFY, 0, 0)
    
    ' Set Property
    If lngIsDirty = 0 Then
        ' False
        blnIsDirty = False
    Else
        ' True
        blnIsDirty = True
    End If
    
    ' Return Value
    GetIsDirty = blnIsDirty
End Function

Private Function GetCaretPosition() As Long
    Dim lngCaretPos As Long
    Dim lngSelStart As Long
    Dim lngSelLen As Long
    Dim lngSelected As Long
    Dim lngLineIndex As Long
    
    ' Grab the Selected Position
    With txtTextBox
        lngSelStart = .SelStart
        lngSelLen = .SelLength
    End With
    
    lngSelected = lngSelStart + lngSelLen
    
    ' Get LINEINDEX
    lngLineIndex = GetLineIndex(GetLineNumber)
    
    ' Calulate Position
    lngCaretPos = (lngSelected - lngLineIndex) + 1
    
    ' Return value
    GetCaretPosition = lngCaretPos
End Function

' ----------------------------------------
' Private Subroutines
' ----------------------------------------

Private Sub SetSoftBreaks()
    Dim lngSoftBreaks As Long
    
    ' Set value
    If blnSoftBreaks = True Then
        ' Set to True
        lngSoftBreaks = SendMessageBynum(txtTextBox.hwnd, EM_FMTLINES, 1, 0)
    Else
        ' Set to false
        lngSoftBreaks = SendMessageBynum(txtTextBox.hwnd, EM_FMTLINES, 0, 0)
    End If
End Sub

Private Sub SetReadOnly()
    Dim lngReadOnly As Long
    
    ' Set value
    If blnReadOnly = True Then
        ' Set to True
        lngReadOnly = SendMessageBynum(txtTextBox.hwnd, EM_SETREADONLY, 1, 0)
    Else
        ' Set to False
        lngReadOnly = SendMessageBynum(txtTextBox.hwnd, EM_SETREADONLY, 0, 0)
    End If
End Sub
