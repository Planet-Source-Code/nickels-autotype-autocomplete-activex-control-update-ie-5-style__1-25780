VERSION 5.00
Begin VB.UserControl ctlAutoType 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ListBox lstFolders 
      Height          =   840
      Left            =   1095
      TabIndex        =   3
      Top             =   2415
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.ListBox lstHistory 
      Height          =   1035
      Left            =   930
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.TextBox txtTypeIn 
      Height          =   315
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.ComboBox cmbTypeIn 
      Height          =   315
      Left            =   15
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   15
      Width           =   3705
   End
End
Attribute VB_Name = "ctlAutoType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'************************************************************************
'AutoType Control v2.0 Copyright 1999 By NeoText
'
'
'Support:
'   nick@neotextsoftware.com
'   support@neotextsoftware.com
'
'   http://www.neotextsoftware.com
'
'
'Terms of Agreement:
' By using this source code, you agree to the following terms...
'  1) You may use this source code in personal projects and may compile
'     it into an .exe/.dll/.ocx and distribute it in binary format
'     freely and with no charge.
'  2) You MAY NOT redistribute this source code (for example to a
'     web site) without written permission from the original author.
'     Failure to do so is a violation of copyright laws.
'  3) You may link to this code from another website, provided it
'     is not wrapped in a frame.
'  4) The author of this code may have retained certain additional
'     copyright rights.If so, this is indicated in the author's
'     description.
'************************************************************************


Enum AutoCompleteStyles
    sNone = 0
    sListComplete = 1
    sTextComplete = 2
End Enum


Private Const Default_Text = "AutoType Control (http://www.neotextsoftware.com)"


Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_HIDE = 0


Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long      ' Get the window co-ordinates in a RECT structure
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type


Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Const LB_FINDSTRING As Long = &H18F
Private Const CB_FINDSTRING As Long = &H14C
Private Const CB_SHOWDROPDOWN As Long = &H14F
Private Const CB_GETCOUNT = &H146
Private Const CB_GETCURSEL = &H147
Private Const CB_GETEDITSEL = &H140
Private Const CB_SELECTSTRING = &H14D
Private Const CB_SETCURSEL = &H14E
Private Const CB_SETEDITSEL = &H142


Private cancelTracking As Boolean

Private isEnabled As Boolean
Private myCompleteStyle As Integer
Private myTracking As Integer
Private maxHistory As Integer
Private isHistoryVisible As Boolean
Private myToolTipText As String
Private isURLEnabled As Boolean
Private isPathEnabled As Boolean

Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)


Private isHooked As Boolean
Private parentHWnd As Long
Private CurrentSelection As Long

Private Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)

    Dim lFlag As Long
    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hWnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW

End Sub

Private Function GetURLText(ByVal pText As String, Optional ByRef pLen As Integer = 0)
    If Len(pText) <= 6 Then
        If LCase(pText) = Left("http://", Len(pText)) Or LCase(pText) = Left("ftp://", Len(pText)) Then
            GetURLText = pText
        Else
            GetURLText = "http://" & pText
            If pLen > 0 Then
                pLen = pLen + 7
            End If
        End If
    Else
        If LCase(Left(pText, 7)) <> "http://" And LCase(Left(pText, 6)) <> "ftp://" Then
            GetURLText = "http://" & pText
            If pLen > 0 Then
                pLen = pLen + 7
            End If
        Else
            GetURLText = pText
        End If
    End If
End Function

Private Sub TrackText()
On Error GoTo catch
'    If Not cancelTracking Then
        cancelTracking = True
        Dim srhText
        Dim srhLen As Integer
        Dim lstIndex As Long
        If isHistoryVisible Then
            Set srhText = cmbTypeIn
        Else
            Set srhText = txtTypeIn
        End If
        srhLen = Len(srhText.Text)
        If (srhText <> "") And (srhLen > 1) Then
            Dim newText As String
            Select Case myCompleteStyle
                Case sTextComplete
                    If isURLEnabled Then
                        newText = GetURLText(srhText.Text, srhLen)
                    Else
                        newText = srhText.Text
                    End If
                    lstIndex = SendMessageStr(lstHistory.hWnd, LB_FINDSTRING, -1, newText)
                    If lstIndex > -1 Then
                        srhText.Text = lstHistory.List(lstIndex)
                        srhText.SelStart = srhLen
                        srhText.SelLength = Len(lstHistory.List(lstIndex)) - srhLen
                    End If
                Case sListComplete
                    Dim lstUse
                    If isURLEnabled Then
                        Dim fso, pFolder
                        Set fso = CreateObject("Scripting.FileSystemObject")
                        If InStr(srhText.Text, "\") > 0 Then
                            pFolder = Left(srhText.Text, InStrRev(srhText.Text, "\"))
                        Else
                            pFolder = srhText.Text
                        End If
                        If fso.FolderExists(pFolder) Then
                            lstFolders.Clear
                            Dim f, fItem
                            Set f = fso.getfolder(pFolder)
                            For Each fItem In f.subfolders
                                lstFolders.AddItem fItem
                            Next
                            For Each fItem In f.Files
                                lstFolders.AddItem fItem
                            Next
                    
                            Set lstUse = lstFolders
                            newText = srhText.Text
                        Else
                            Set lstUse = lstHistory
                            newText = GetURLText(srhText.Text)
                        End If
                    Else
                        Set lstUse = lstHistory
                        newText = srhText.Text
                    End If
                    
                
                    lstIndex = SendMessageStr(lstUse.hWnd, LB_FINDSTRING, -1, newText)
                    If lstIndex > -1 Then
                        ResizeComboList
                        SendMessageLong frmCombo.hWnd, SW_SHOWNOACTIVATE, False, 0&
                        AlwaysOnTop frmCombo, True
                        frmCombo.lstMatch.Clear
                        Set frmCombo.srhText = Nothing
                        Set frmCombo.srhText = srhText
                        Dim oldIndex As Long
                        lstIndex = -1
                        Do
                            oldIndex = lstIndex
                            lstIndex = SendMessageStr(lstUse.hWnd, LB_FINDSTRING, oldIndex, newText)
                            If (lstIndex > -1) And Not (lstIndex <= oldIndex) Then
                                frmCombo.lstMatch.AddItem lstUse.List(lstIndex)
                            End If
                        Loop Until lstIndex = -1 Or lstIndex = lstUse.ListCount - 1 Or (lstIndex <= oldIndex)
                    Else
                        frmCombo.Visible = False
                    End If
                    
                    CurrentSelection = -1
            End Select
        End If
        cancelTracking = False
'    End If
Exit Sub
catch:
    Err.Clear
End Sub



Public Function ShowHistory()
    If Not isHistoryVisible Then
        Dim r As Long
        r = SendMessageLong(cmbTypeIn.hWnd, CB_SHOWDROPDOWN, True, 0)
    End If
End Function
Public Function HideHistory()
    If Not isHistoryVisible Then
        Dim r As Long
        r = SendMessageLong(cmbTypeIn.hWnd, CB_SHOWDROPDOWN, False, 0)
    End If
End Function

Public Sub SetFocus()
On Error Resume Next
    If isHistoryVisible Then
        SetFocusAPI cmbTypeIn.hWnd
    Else
        SetFocusAPI txtTypeIn.hWnd
    End If
Err.Clear
End Sub


Public Property Get SelStart() As Integer
    If isHistoryVisible Then
        SelStart = cmbTypeIn.SelStart
    Else
        SelStart = txtTypeIn.SelStart
    End If
End Property
Public Property Let SelStart(ByVal newValue As Integer)
    If isHistoryVisible Then
        cmbTypeIn.SelStart = newValue
    Else
        txtTypeIn.SelStart = newValue
    End If
End Property


Public Property Get SelLength() As Integer
    If isHistoryVisible Then
        SelLength = cmbTypeIn.SelLength
    Else
        SelLength = txtTypeIn.SelLength
    End If
End Property
Public Property Let SelLength(ByVal newValue As Integer)
    If isHistoryVisible Then
        cmbTypeIn.SelLength = newValue
    Else
        txtTypeIn.SelLength = newValue
    End If
End Property


Public Sub SetToList(ByVal lstBox As Variant)
    Me.ClearHistory
    Dim cnt As Integer
    If lstBox.ListCount > 0 Then
        For cnt = 0 To lstBox.ListCount - 1
            Me.AddItem lstBox.List(cnt)
        Next
    End If
End Sub


Public Sub AddItem(ByVal lstText As String)
    If lstHistory.ListCount < maxHistory - 1 Then
        Dim lstIndex As Integer
        lstIndex = SendMessageStr(lstHistory.hWnd, LB_FINDSTRING, -1, lstText)
        If lstIndex = -1 Then
            lstHistory.AddItem lstText
            cmbTypeIn.AddItem lstText
        End If
    End If
End Sub
Public Sub RemoveItem(ByVal lstIndex As Integer)
    lstHistory.RemoveItem lstIndex
    cmbTypeIn.RemoveItem lstIndex
End Sub
Public Sub ClearHistory()
    Do Until lstHistory.ListCount <= 0
        lstHistory.RemoveItem 0
    Loop
    Do Until cmbTypeIn.ListCount <= 0
        cmbTypeIn.RemoveItem 0
    Loop
End Sub


Public Property Get ListCount() As Integer
    ListCount = lstHistory.ListCount
End Property


Public Property Get HistorySize() As Integer
    HistorySize = maxHistory
End Property
Public Property Let HistorySize(ByVal newValue As Integer)
    maxHistory = newValue
End Property


Public Property Get HistoryVisible() As Boolean
    HistoryVisible = isHistoryVisible
End Property
Public Property Let HistoryVisible(ByVal newValue As Boolean)
    isHistoryVisible = newValue
    cmbTypeIn.Visible = newValue
    txtTypeIn.Visible = Not newValue
End Property


Public Property Get ToolTipText() As String
    ToolTipText = myToolTipText
End Property
Public Property Let ToolTipText(ByVal newValue As String)
    myToolTipText = newValue
    cmbTypeIn.ToolTipText = myToolTipText
    txtTypeIn.ToolTipText = myToolTipText
End Property


Public Property Get Enabled() As Boolean
    Enabled = isEnabled
End Property
Public Property Let Enabled(ByVal newValue As Boolean)
    isEnabled = newValue
    cmbTypeIn.Enabled = isEnabled
    txtTypeIn.Enabled = isEnabled
End Property


Public Property Get Text() As String
    If isHistoryVisible Then
        Text = cmbTypeIn.Text
    Else
        Text = txtTypeIn.Text
    End If
End Property
Public Property Let Text(ByVal newText As String)
    cmbTypeIn.Text = newText
    txtTypeIn.Text = newText
End Property


Public Property Get CompleteStyle() As AutoCompleteStyles
    CompleteStyle = myCompleteStyle
End Property
Public Property Let CompleteStyle(ByVal newValue As AutoCompleteStyles)
    myCompleteStyle = newValue
End Property

Public Property Get URLEnabled() As Boolean
    URLEnabled = isURLEnabled
End Property
Public Property Let URLEnabled(ByVal newValue As Boolean)
    isURLEnabled = newValue
End Property

Public Property Get PathEnabled() As Boolean
    PathEnabled = isPathEnabled
End Property
Public Property Let PathEnabled(ByVal newValue As Boolean)
    isPathEnabled = newValue
End Property





'*************************************************************************
'*************************************************************************
'*************************************************************************
'*************************************************************************
'*************************************************************************
Private Sub cmbTypeIn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 8 Then
    
    End If
    If isHistoryVisible Then
        txtObject_KeyDown KeyCode, Shift
    End If
End Sub
Private Sub cmbTypeIn_KeyPress(KeyAscii As Integer)
    If isHistoryVisible Then
        txtObject_KeyPress KeyAscii
    End If
End Sub
Private Sub cmbTypeIn_KeyUp(KeyCode As Integer, Shift As Integer)
    If isHistoryVisible Then
        txtObject_KeyUp KeyCode, Shift
    End If
End Sub
Private Sub cmbTypeIn_Change()
    If isHistoryVisible Then
        TrackText
        txtObject_Change
    End If
End Sub
Private Sub cmbTypeIn_Click()
    If isHistoryVisible Then
        txtObject_Click
    End If
End Sub
Private Sub cmbTypeIn_DblClick()
    If isHistoryVisible Then
        txtObject_DblClick
    End If
End Sub
'*************************************************************************
'*************************************************************************
'*************************************************************************
Private Sub txtTypeIn_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not isHistoryVisible Then
        txtObject_KeyDown KeyCode, Shift
    End If
End Sub
Private Sub txtTypeIn_KeyPress(KeyAscii As Integer)
    If Not isHistoryVisible Then
        txtObject_KeyPress KeyAscii
    End If
End Sub
Private Sub txtTypeIn_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not isHistoryVisible Then
        txtObject_KeyUp KeyCode, Shift
    End If
End Sub
Private Sub txtTypeIn_Change()
    If Not isHistoryVisible Then
        If Not cancelTracking Then TrackText
        txtObject_Change
    End If
End Sub
Private Sub txtTypeIn_Click()
    If Not isHistoryVisible Then
        txtObject_Click
    End If
End Sub
Private Sub txtTypeIn_DblClick()
    If Not isHistoryVisible Then
        txtObject_DblClick
    End If
End Sub
Private Sub txtTypeIn_GotFocus()
    If Not isHistoryVisible Then
        txtObject_GotFocus
    End If
End Sub
'*************************************************************************
'*************************************************************************
'*************************************************************************
Private Sub txtObject_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Or KeyCode = 8 Then
        cancelTracking = True
    Else
        cancelTracking = False
    End If
    If frmCombo.lstMatch.ListCount > 0 And frmCombo.Visible = True Then
        Select Case KeyCode
            Case 38
                KeyCode = 0
            Case 40
                SetFocusAPI frmCombo.lstMatch.hWnd
                frmCombo.lstMatch.ListIndex = 0
                KeyCode = 0
            Case 13
                frmCombo.Visible = False
        End Select
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub txtObject_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub txtObject_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub txtObject_Change()
    RaiseEvent Change
End Sub
Private Sub txtObject_Click()
    RaiseEvent Click
End Sub
Private Sub txtObject_DblClick()
    RaiseEvent DblClick
End Sub
Private Sub txtObject_GotFocus()
    txtTypeIn.SelStart = 0
    txtTypeIn.SelLength = Len(txtTypeIn.Text)
End Sub
'*************************************************************************
'*************************************************************************
'*************************************************************************
'*************************************************************************
'*************************************************************************


Private Sub UserControl_Initialize()
    cancelTracking = False
    Load frmCombo

End Sub
Private Sub UserControl_InitProperties()
    Me.CompleteStyle = sListComplete
    Me.HistorySize = 20
    Me.HistoryVisible = True
    Me.Enabled = True
    Me.URLEnabled = True
    Me.PathEnabled = True
    Me.ToolTipText = ""
    Me.Text = Default_Text
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        myCompleteStyle = .ReadProperty("CompleteStyle", sListComplete)
        Me.CompleteStyle = myCompleteStyle
        maxHistory = .ReadProperty("HistorySize", 20)
        Me.HistorySize = maxHistory
        isHistoryVisible = .ReadProperty("HistoryVisible", True)
        Me.HistoryVisible = isHistoryVisible
        isEnabled = .ReadProperty("Enabled", True)
        Me.Enabled = isEnabled
        isURLEnabled = .ReadProperty("URLEnabled", True)
        Me.URLEnabled = isURLEnabled
        isPathEnabled = .ReadProperty("PathEnabled", True)
        Me.PathEnabled = isPathEnabled
        myToolTipText = .ReadProperty("ToolTipText", "")
        Me.ToolTipText = myToolTipText
        cmbTypeIn.Text = .ReadProperty("Text", Default_Text)
        txtTypeIn.Text = .ReadProperty("Text", Default_Text)
        
    End With
End Sub

Private Sub UserControl_Show()
    parentHWnd = UserControl.Parent.hWnd
    Hook UserControl.Parent.hWnd, isHooked
End Sub

Private Sub UserControl_Hide()
    UnHook parentHWnd, isHooked
End Sub

Private Sub UserControl_Terminate()
    Unload frmCombo
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "CompleteStyle", myCompleteStyle, sListComplete
        .WriteProperty "HistorySize", maxHistory, 20
        .WriteProperty "HistoryVisible", isHistoryVisible, True
        .WriteProperty "Enabled", isEnabled, True
        .WriteProperty "URLEnabled", isURLEnabled, True
        .WriteProperty "PathEnabled", isPathEnabled, True
        .WriteProperty "ToolTipText", myToolTipText, ""
        If isHistoryVisible Then
            .WriteProperty "Text", cmbTypeIn.Text, Default_Text
        Else
            .WriteProperty "Text", txtTypeIn.Text, Default_Text
        End If
    End With
End Sub
Private Sub UserControl_Resize()
    cmbTypeIn.Top = 15
    cmbTypeIn.Left = 15
    txtTypeIn.Top = 15
    txtTypeIn.Left = 15
    If Height <> 345 Then Height = 345
    cmbTypeIn.Width = UserControl.Width - 30
    txtTypeIn.Width = UserControl.Width - 30
    
    ResizeComboList
End Sub

Private Sub ResizeComboList()
    On Error Resume Next
    Dim rctCombo As RECT
    GetWindowRect cmbTypeIn.hWnd, rctCombo
    
    frmCombo.Top = (rctCombo.Bottom * Screen.TwipsPerPixelY)
    frmCombo.Left = (rctCombo.Left * Screen.TwipsPerPixelX)
    frmCombo.Width = ((rctCombo.Right - rctCombo.Left) * Screen.TwipsPerPixelX)
    frmCombo.ResizeBox
    Err.Clear
End Sub






