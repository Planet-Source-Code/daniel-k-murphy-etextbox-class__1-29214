VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUnSelect 
      Caption         =   "UnSelect Text"
      Height          =   495
      Left            =   5400
      TabIndex        =   20
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdDirty 
      Caption         =   "Is Dirty?"
      Height          =   495
      Left            =   5400
      TabIndex        =   19
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCanUndo 
      Caption         =   "Can Undo?"
      Height          =   495
      Left            =   5400
      TabIndex        =   18
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdClearUndoBuffer 
      Caption         =   "Clear Undo Buffer"
      Height          =   495
      Left            =   5400
      TabIndex        =   17
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "Undo"
      Height          =   495
      Left            =   4080
      TabIndex        =   16
      Top             =   4320
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   2760
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSaveText 
      Caption         =   "Save"
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoadText 
      Caption         =   "Load"
      Height          =   495
      Left            =   4080
      TabIndex        =   14
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdReadOnly 
      Caption         =   "Read Only"
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdCaretPos 
      Caption         =   "Caret Pos"
      Height          =   495
      Left            =   2760
      TabIndex        =   12
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdLineIndex 
      Caption         =   "Line Index"
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdLineInfo 
      Caption         =   "Line Info"
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdLineLength 
      Caption         =   "Line Length"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdWordCount 
      Caption         =   "Word Count"
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCharCount 
      Caption         =   "Char Count"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdBottomLine 
      Caption         =   "Bottom Line"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdTopLine 
      Caption         =   "Top Line"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdNumLines 
      Caption         =   "Number Lines"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdVisibleLines 
      Caption         =   "Visible Lines"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetLine 
      Caption         =   "Get Line"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00000000&
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declare the Class
Private eTextBox As clsETextBox

Private Sub cmdBottomLine_Click()
    With eTextBox
        MsgBox "Bottom Line Used is " + Trim(Str(.BottomLine))
    End With
End Sub

Private Sub cmdCanUndo_Click()
    With eTextBox
        If .CanUndo Then
            MsgBox "Undo is possible!"
        Else
            MsgBox "Cannot Undo!"
        End If
    End With
End Sub

Private Sub cmdCaretPos_Click()
    With eTextBox
        MsgBox "Caret Pos: " + Trim(Str(.CaretPos))
    End With
End Sub

Private Sub cmdCharCount_Click()
    With eTextBox
        MsgBox "Total number of Characters: " + Trim(Str(.CharacterCount))
    End With
End Sub

Private Sub cmdClear_Click()
    With eTextBox
        .Clear
    End With
End Sub

Private Sub cmdClearUndoBuffer_Click()
    With eTextBox
        .ClearUndoBuffer
    End With
End Sub

Private Sub cmdDirty_Click()
    With eTextBox
        If .IsDirty Then
            MsgBox "TextBox has changes"
        Else
            MsgBox "TextBox was not changed"
        End If
    End With
End Sub

Private Sub cmdGetLine_Click()
    With eTextBox
        MsgBox "You are on line # " + Trim(Str(.LineNumber))
    End With
End Sub

Private Sub cmdLineIndex_Click()
    With eTextBox
        MsgBox "Line Index is: " + Trim(.LineIndex(CLng(.LineNumber)))
    End With
End Sub

Private Sub cmdLineInfo_Click()
    With eTextBox
        MsgBox "Line Info is: " + Trim(.LineData(CLng(.LineNumber)))
    End With
End Sub

Private Sub cmdLoadText_Click()
    Dim strFileName As String
    
    strFileName = cdlg.FileName
    
    cdlg.ShowOpen
    
    If strFileName = cdlg.FileName Then
        If MsgBox("This will restore to the original!", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    strFileName = cdlg.FileName
    
    With eTextBox
        .LoadTEXT strFileName
    End With
End Sub

Private Sub cmdNumLines_Click()
    With eTextBox
        MsgBox "There are " + Trim(Str(.LineCount)) + " Lines"
    End With
End Sub

Private Sub cmdReadOnly_Click()
    With eTextBox
        If cmdReadOnly.Caption = "Read Only" Then
            .ReadOnly = True
            cmdReadOnly.Caption = "Write"
        Else
            .ReadOnly = False
            cmdReadOnly.Caption = "Read Only"
        End If
    End With
End Sub

Private Sub cmdSaveText_Click()
    Dim strFileName As String
    
    cdlg.ShowSave
    
    strFileName = cdlg.FileName
    
    With eTextBox
        .SaveTEXT strFileName
    End With
End Sub

Private Sub cmdTopLine_Click()
    With eTextBox
        MsgBox "Top Visible Line # " + Trim(Str(.TopLine))
    End With
End Sub

Private Sub cmdUndo_Click()
    With eTextBox
        .Undo
    End With
End Sub

Private Sub cmdUnSelect_Click()
    With eTextBox
        .UnSelect
    End With
End Sub

Private Sub cmdVisibleLines_Click()
    With eTextBox
        MsgBox "There are " + Trim(Str(.VisibleLines)) + " Lines Visible"
    End With
End Sub

Private Sub cmdWordCount_Click()
    With eTextBox
        MsgBox "Total Number of Words: " + Trim(Str(.WordCount))
    End With
End Sub

Private Sub cmdLineLength_Click()
    Dim lngLineNumber As Long
    Dim lngLineLength As Long
    
    With eTextBox
        ' Get Line Number
        lngLineNumber = .LineNumber
        
        ' Get Line Length
        lngLineLength = .LineLength(lngLineNumber)
        
        ' Show Line Length
        MsgBox "Line Length is " + Trim(Str(lngLineLength)) + " Characters"
    End With
End Sub

Private Sub Form_Load()
    ' Capture the Title of the program
    Caption = App.Title
    
    ' Assign the textbox
    Set eTextBox = New clsETextBox
    
    With eTextBox
        Set .eTextBox = Text1
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Free the class
    Set eTextBox = Nothing
End Sub
