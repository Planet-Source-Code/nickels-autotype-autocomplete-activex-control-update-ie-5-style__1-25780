'AutoType Control v1.0 By NeoText
'
'Support:
'	flash@quakeclan.net
'	neotext@quakeclan.net
'	neotext@email.com
'	
'	http://www.quakeclan.net/neotext/
'
'
'Properties
'
'    .CompleteStyle = Integer
'        Sets if the auto type control should display a list.
'        Style Values:
'            sNone         = 0 Disables auto complete
'            sListComplete = 1 List pops up with matches to what you type.
'            sTextComplete = 2 No List pops up, the closes match is put in as you type.
'
'    .HistoryVisible = Boolean
'        Sets if the History drop down list should be visible.
'        The history list is seperate from the AutoComplete List.
'
'    .URLEnabled = Boolean
'        Sets if the auto complete should recognize URLs
'
'    .PathEnabled = Boolean
'        Sets if the auto complete should include the local file system.
'
'    .Enabled = Boolean
'        Enables/Disables the control. (standard TextBox Property)
'
'    .ToolTipText = String
'        Sets the ToolTipText of the control.
'
'    .HistorySize = Integer
'        Sets the max size of history entries that are allowed.
'
'    .ListCount = Integer
'        Read-only, returns the number of items in the History.
'
'    .Text = String
'        Sets the text of the control. (standard TextBox Property)
'
'    .SelStart = Integer
'        Sets the start position of the text selection. (standard TextBox Property)
'
'    .SelLength = Integer
'        Sets the length of the text selection. (standard TextBox Property)
'
'Methods
'
'    .SetToList (lstBox as ListBoxControl)
'        Sets up the History list to lstBox's items, limited by HistorySize.
'
'    .AddItem (ItemText as String)
'        Adds ItemText to the History list, limited by HistorySize.
'        Doesn't add duplicate values, checks for duplicates with .CaseSensitive
'
'    .RemoveItem (ItemIndex as Integer)
'        Removes ItemIndex from the History list.
'
'    .ClearHistory
'        Clears the History list.
'
'    .ShowHistory
'        Scrolls down the drop down list.  Only works if HistoryVisible = True
'
'    .HideHistory
'        Scrolls up the drop down list.  Only works if HistoryVisible = True
'
'    .SetFocus
'        Sets focus to the control
'
'Events - (Standard TextBox Events)
'    Change()
'    Click()
'    DblClick()
'    GotFocus()
'    LostFocus()
'    KeyDown(KeyCode As Integer, Shift As Integer)
'    KeyPress(KeyAscii As Integer)
'    KeyUp(KeyCode As Integer, Shift As Integer)
'
'
'Consts
'    CompleteType Constants
'        sNone = 0
'        sListComplete = 1
'        sTextComplete = 2
