VERSION 5.00
Begin VB.Form frmReplace 
   Caption         =   "Find/Replace"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   LinkTopic       =   "Form3"
   ScaleHeight     =   2250
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   1350
      Width           =   1335
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find First"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   900
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Replace"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Replace"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Find"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy All"
      End
   End
End
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author:   dzzie@yahoo.com
'Site:     http://sandsprite.com
'this form is no longer used universally, scivb_lite has a copy of this form built in use that for txtjs


Public active_object As RichTextBox
Dim lastkey As Integer
Dim lastIndex As Long
Dim lastsearch As String

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_SHOWWINDOW = &H40

Dim selli As ListItem



Private Sub cmdfind_Click()
    
    On Error Resume Next
    
     
        f = Text1
     
    
    lastsearch = f
    
    Dim compare As VbCompareMethod
    
     
        compare = vbTextCompare
    
    
    x = InStr(1, active_object.Text, lastsearch, compare)
    If x > 0 Then
        lastIndex = x + 2
        active_object.SelStart = x - 1
        active_object.SelLength = Len(lastsearch)
    Else
        lastIndex = 1
    End If
    
End Sub


Private Sub cmdFindNext_Click()
    
    On Error Resume Next
    
    
        f = Text1
    
    
    If lastsearch <> f Then
        cmdfind_Click
        Exit Sub
    End If
    
    If lastIndex >= Len(active_object.Text) Then
        MsgBox "Reached End of text no more matches", vbInformation
        Exit Sub
    End If
    
    Dim compare As VbCompareMethod
    
     
        compare = vbTextCompare
     
    
    x = InStr(lastIndex, active_object.Text, lastsearch, compare)
    
    If x + 2 = lastIndex Or x < 1 Then
        MsgBox "No more matches found", vbInformation
        Exit Sub
    Else
        lastIndex = x + 2
        active_object.SelStart = x - 1
        active_object.SelLength = Len(lastsearch)
    End If
    
    
End Sub

Private Sub Command1_Click()
    
    On Error Resume Next
    
    
        f = Text1
     
    
     
        r = Text2
     
    
    Dim compare As VbCompareMethod
    
     
        compare = vbTextCompare
     
    
    Dim curLine As Long
    
     
        active_object.Text = Replace(active_object.Text, f, r, , , compare)
     
    
     
    
End Sub

Public Sub LaunchReplaceForm(txtObj As RichTextBox)
    
    Set active_object = txtObj
    
     
    
    Me.Show
    
End Sub




Private Sub Form_Load()
    Me.Icon = Form1.Icon
    'FormPos Me, False
    SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, SWP_SHOWWINDOW
    'Text1 = GetMySetting("lastFind")
    'Text2 = GetMySetting("lastReplace")
    'If GetMySetting("wholeText", "1") = "1" Then Option1.Value = True Else Option2.Value = True
End Sub

'Private Sub Form_Resize()
'    On Error Resume Next
'    lv.Width = Me.Width - lv.Left - 200
'    lv.Height = Me.Height - lv.Top - 300
'    lv.ColumnHeaders(2).Width = lv.Width - lv.ColumnHeaders(2).Left - 200
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    'FormPos Me, False, True
    'SaveMySetting "lastFind", Text1
    'SaveMySetting "lastReplace", Text2
    'SaveMySetting "wholeText", IIf(Option1.Value, "1", "0")
End Sub

 
