Attribute VB_Name = "modSubclass"
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

Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long


Private Const GWL_WNDPROC = -4
Private Const WM_CLOSE = &H10
Private Const WM_SYSCOMMAND = &H112
Private Const WM_ACTIVATE = &H6
Private Const WM_CHILDACTIVATE = &H22
Private Const WM_KILLFOCUS = &H8
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_NCACTIVATE = &H86
Private Const WM_WINDOWPOSCHANGED = &H47
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwnewlong As Long) As Long
Private OldWinProc As Long
Public Function WindowMessageProc(ByVal inHWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case Msg
        Case WM_CLOSE
        Case WM_ACTIVATE
            If wParam = 0 And lParam = 0 Then
                frmCombo.Visible = False
            End If
        Case 13
            If wParam = 510 And lParam = 1240124 Then
                frmCombo.Visible = False
            End If
        Case 4110
            frmCombo.Visible = False
        Case 533
            frmCombo.Visible = False
        Case Else
            'Debug.Print "Message: " & Msg & " wParam: " & wParam & " lParam: " & lParam
    End Select
    WindowMessageProc = CallWindowProc(OldWinProc, inHWnd, Msg, wParam, lParam)
End Function
Public Function Hook(ByVal hWnd As Long, ByRef Hooked As Boolean)
    If OldWinProc = 0 Then
        OldWinProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowMessageProc)
        Hooked = True
    Else
        Hooked = False
    End If
End Function
Public Function UnHook(ByVal hWnd As Long, ByVal Hooked As Boolean)
    If Hooked Then SetWindowLong hWnd, GWL_WNDPROC, OldWinProc
End Function





