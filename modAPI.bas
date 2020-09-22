Attribute VB_Name = "modAPI"
Option Explicit

' ----------------------------------
' Declare the API Functions
' ----------------------------------
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageBynum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageBystring Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDC As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

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
Private lngLineNumber As Long

Public Property Set eTextBox(eText As TextBox)
    Set txtTextBox = eText
End Property

Public Property Get LineNumber() As Double
    If txtTextBox = Null Then
        ' No textbox is set
        ' Return -1
        lngLineNumber = -1
    Else
        ' Get the Position
        lngLineNumber = GetLineNumber
    End If
    
    ' Return Value
    LineNumber = lngLineNumber
End Property

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
