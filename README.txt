-------------------------------------------------
		eTextBox Class
-------------------------------------------------

Written By: Daniel K Murphy

EMail: daniel_k_murphy@hotmail.com

Other Contributors: Christopher D Lucas
			WordCount
			CharacterCount

-----------------------
Purpose
-----------------------

Enhance the TextBox control so that the following capabilites are possible:

-----------------------
Properties
-----------------------
BottomLine -- Returns (Double value) the Last Visible line in the textbox.  If there are less actual used lines than there are total lines visible in the textbox, then the last line containing data is returned. (-1) is returned if no textbox is set.

CanUndo -- Returns (Boolean) if Undo is possible.  True if possible, False in error or not possible.

CaretPos -- Returns (Double Value) the position in the line. (-1) is returned if no text box is set.

CharacterCount -- Returns (Double Value) the total number of characters in the textbox. (-1) is returned if no textbox is set.

eTextBox -- Sets the textbox to work with.  This  MUST be done before any property or function is able to be used. (ex: Set eTextBox.eTextBox = Text1, assuming you declared eTextBox as New clsETextBox)

IsDirty -- Returns (Boolean) if there was a change to the textbox. False is returned if not textbox is set or no change was made.

LineCount -- Returns (Double Value) the total number of lines in the textbox.  (-1) is returned if no textbox is set.

LineNumber -- Returns (Double Value) the current line number at which the Caret is positioned. This is a Zero Based property, so it will ALWAYS be one less than LineCount if on the last line in the textbox. (-1) is returned if no textbox is set.

ReadOnly -- Returns/Sets (Boolean) if the textbox is read-only.  True locks the textbox, but DOES NOT grey out the text.

SoftBreaks -- Sets (Boolean) if the the returning string of text forces a Hard-Break where soft-breaks occur on the Viewable Textbox control.  False forces Soft-Breaks to return as Hard-Breaks.

TopLine -- Returns (Double Value) the first visible line in the textbox.  If you are 3 lines  scrolled down from Line 0, then it will return 2.  This is a Zero Based property.  (-1) is returned if no textbox is set.

VisibleLines -- Returns (Double Value) the total number of visible lines in a textbox.  (-1) is returned if no textbox is set.

WordCount -- Returns (Double Value) the total number of words in the textbox.  (-1) is returned if no textbox is set.

-----------------------
Functions
-----------------------
LineData(LineNum as Long) -- Returns (String Value) the data on a specific line in a textbox.  "ERROR" is returned if no textbox is set.

LineIndex(LineNum as Long) -- Returns (Double Value) the Position number of the First character on a specific line in a textbox.  (-1) is returned if no textbox is set.

LineLength(LineNum as Long) -- Returns (Double Value) the length of a specific line in a textbox.  (-1) is returned if no textbox is set.

Undo -- Returns (Boolean) True if successful, False if unsuccessful.  If successful, it restores the text before the last changes made.

-----------------------
Subroutines
-----------------------
Clear -- Removes all text in the TextBox.

ClearUndoBuffer -- Removes the Undo buffer.  Once executed, Undo is no possible until the buffer is filled again.

LoadTEXT -- Loads the text from a file into the TextBox.

SaveTEXT -- Saves the text in a file from the TextBox.

UnSelect -- Clears the Selected text.
