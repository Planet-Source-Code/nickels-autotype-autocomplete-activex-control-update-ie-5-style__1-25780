VERSION 5.00
Begin VB.Form frmCombo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   990
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox lstMatch 
      CausesValidation=   0   'False
      Height          =   1035
      ItemData        =   "frmCombo.frx":0000
      Left            =   -30
      List            =   "frmCombo.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   -30
      Width           =   2865
   End
End
Attribute VB_Name = "frmCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Public srhText

Public Sub ResizeBox()
    On Error Resume Next
    lstMatch.Top = -30
    lstMatch.Left = -30
    lstMatch.Width = Me.ScaleWidth + (60)
    Err.Clear
End Sub

Private Sub SetText(ByVal KeyAscii As Integer)
    If lstMatch.ListIndex >= 0 Then
        SetFocusAPI srhText.hWnd
        Me.Visible = False
        srhText.Text = lstMatch.List(lstMatch.ListIndex)
        srhText.SelStart = Len(srhText.Text)
        If KeyAscii > 0 Then
            SendMessageLong srhText.hWnd, WM_CHAR, KeyAscii, 0
        End If
    End If
End Sub

Private Sub lstMatch_DblClick()
    SetText 0
    'SendMessageLong srhText.hWnd, WM_CHAR, 13, 0
End Sub

Private Sub lstMatch_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case 38
            If frmCombo.lstMatch.ListIndex = 0 Then
                SetFocusAPI srhText.hWnd
                srhText.SelStart = Len(srhText.Text)
                lstMatch.ListIndex = -1
            End If
        Case 13
            SetText 0
            'SendMessageLong srhText.hWnd, WM_CHAR, 13, 0
    End Select

End Sub

Private Sub lstMatch_KeyPress(KeyAscii As Integer)
    SetText KeyAscii
End Sub
