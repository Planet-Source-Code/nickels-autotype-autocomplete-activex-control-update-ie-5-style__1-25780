VERSION 5.00
Object = "*\AAutoType.vbp"
Begin VB.Form frmExample 
   Caption         =   "AutoType Example Application"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin AutoType.ctlAutoType atExample1 
      Height          =   345
      Left            =   300
      TabIndex        =   3
      Top             =   405
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   609
   End
   Begin AutoType.ctlAutoType atExample2 
      Height          =   345
      Left            =   3000
      TabIndex        =   2
      Top             =   1185
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   609
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   195
      TabIndex        =   0
      Top             =   1530
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label Label3 
      Caption         =   "Example of ListComplete"
      Height          =   270
      Left            =   3075
      TabIndex        =   5
      Top             =   900
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Example of AutoComplete"
      Height          =   300
      Left            =   345
      TabIndex        =   4
      Top             =   105
      Width           =   2430
   End
   Begin VB.Label Label1 
      Caption         =   "List used for Example 2"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   1230
      Visible         =   0   'False
      Width           =   2325
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub atExample1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        MsgBox "GO 1"
    End If
End Sub

Private Sub atExample2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        MsgBox "GO 2"
    End If
End Sub

Private Sub FillList(ByRef myList)
    With myList
        .AddItem "http://www.ebay.com"
        .AddItem "http://www.amazon.com"
        .AddItem "http://www.altavista.com"
        .AddItem "http://www.microsoft.com"
        .AddItem "http://www.msn.com"
        .AddItem "http://msdn.microsoft.com"
        .AddItem "http://www.planet-source-code.com"
        .AddItem "http://www.neotextsoftware.com"
        .AddItem "http://www.yahoo.com"
        .AddItem "http://www.lycos.com"
        .AddItem "http://www.excite.com"
        .AddItem "http://www.hotmail.com"
        .AddItem "ftp://www.ebay.com"
        .AddItem "ftp://www.amazon.com"
        .AddItem "ftp://www.altavista.com"
        .AddItem "ftp://www.microsoft.com"
        .AddItem "ftp://www.msn.com"
        .AddItem "ftp://msdn.microsoft.com"
        .AddItem "ftp://www.planet-source-code.com"
        .AddItem "ftp://www.neotextsoftware.com"
        .AddItem "ftp://www.yahoo.com"
        .AddItem "ftp://www.lycos.com"
        .AddItem "ftp://www.excite.com"
        .AddItem "ftp://www.hotmail.com"
    End With
End Sub

Private Sub Form_Load()

'############################################
'Example 1:
'   This example uses the AutoType
'   controls .AddItem method to make
'   the list of items in the control.


'add items to the AutoType list directly
FillList atExample1

'set it so the user can not view history by clickin the down arrow
atExample1.historyvisible = False

'set it so the auto complete finishes your typeing instead of poping up a list
atExample1.completeStyle = sTextComplete

'URLEnabled makes it so you don't have to type "http://" and the autocomplete still works
atExample1.urlenabled = False

'PathEnabled makes it so when you type in a folder, auto complete will search the filesystem for you
atExample1.pathenabled = False


'############################################







'############################################
'Example 2:
'    This example uses lstExample2 a listbox
'    to set the AutoType item list by using
'    the method .SetToList



'add files to a ListBox on the form
FillList List1

'set the autotype list to the list on this form
atExample2.SetToList List1


'############################################
End Sub
