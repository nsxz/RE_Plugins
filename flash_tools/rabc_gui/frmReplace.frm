VERSION 5.00
Begin VB.Form frmReplace 
   Caption         =   "Find/Replace"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11550
   LinkTopic       =   "Form3"
   ScaleHeight     =   2265
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFindAll 
      Caption         =   "Find All"
      Height          =   375
      Left            =   3825
      TabIndex        =   8
      Top             =   945
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   5355
      TabIndex        =   7
      Top             =   0
      Width           =   6135
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   2430
      TabIndex        =   6
      Top             =   945
      Width           =   1335
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find First"
      Height          =   375
      Left            =   990
      TabIndex        =   5
      Top             =   945
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Replace"
      Height          =   375
      Left            =   3825
      TabIndex        =   4
      Top             =   1485
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

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_LINESCROLL = &HB6

Dim selli As ListItem



Private Sub cmdfind_Click()
    
    On Error Resume Next
    
    f = Text1
    lastsearch = f
    Me.Caption = "Find / Replace"
    
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


 
Public Sub cmdFindAll_Click()
    
    On Error Resume Next
    Dim txt As String
    Dim line As Long
    Dim editorText As String
    
    If Me.Width < 10440 Then Me.Width = 10440
    List1.Clear
    f = Text1
    
    Dim compare As VbCompareMethod
    compare = vbTextCompare

    lastIndex = 1
    lastsearch = f
    x = 1
    
    If Len(f) = 0 Then Exit Sub
    
    LockWindowUpdate active_object.hwnd
    editorText = active_object.Text
    Do While x > 0
    
        x = InStr(lastIndex, editorText, lastsearch, compare)
    
        If x + 2 = lastIndex Or x < 1 Or x >= Len(editorText) Then
            Exit Do
        Else
            lastIndex = x + 2
            active_object.SelStart = x - 1
            active_object.SelLength = Len(lastsearch)
            line = GetCurrentLine(active_object)
            txt = Replace(Trim(GetLineText(active_object, line)), vbTab, Empty)
            txt = Replace(txt, vbCrLf, Empty)
            txt = Replace(txt, vbLf, Empty)
            While InStr(txt, "  ") > 0
                txt = Replace(txt, "  ", " ")
            Wend
            
            List1.AddItem (line + 1) & ": " & txt
            ScrollToLine active_object, line + 1
             
            
        End If
        
    Loop
    
    LockWindowUpdate 0
    
    If List1.ListCount >= 0 Then
        List1.Selected(0) = True
        List1_Click
    End If
    
    Me.Caption = List1.ListCount & " items found!"
    
End Sub
 

Private Sub cmdFindNext_Click()
    
    On Error Resume Next

    f = Text1
    Me.Caption = "Find / Replace"
     
    If lastsearch <> f Then
        cmdfind_Click
        Exit Sub
    End If
    
    If lastIndex >= Len(active_object.Text) Then
        Me.Caption = "Reached End of text no more matches"
        Exit Sub
    End If
    
    Dim compare As VbCompareMethod
    compare = vbTextCompare
     
    x = InStr(lastIndex, active_object.Text, lastsearch, compare)
    
    If x + 2 = lastIndex Or x < 1 Then
        Me.Caption = "No more matches found"
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

    active_object.Text = Replace(active_object.Text, f, r, , , compare)

End Sub

Public Sub LaunchReplaceForm(txtObj As RichTextBox)
    
    Set active_object = txtObj
    Me.Show
    
End Sub




Private Sub Form_Load()
    'FormPos Me, False
    SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, SWP_SHOWWINDOW
    'Text1 = GetMySetting("lastFind")
    'Text2 = GetMySetting("lastReplace")
    'If GetMySetting("wholeText", "1") = "1" Then Option1.Value = True Else Option2.Value = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Width = Me.Width - List1.Left - 200
    List1.Height = Me.Height - List1.Top - 400
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

 
 Sub ScrollToLine(rtf As RichTextBox, Number As Long)
    Dim curLine As Long, Shift As Long
    curLine = SendMessage(rtf.hwnd, EM_GETFIRSTVISIBLELINE, 0&, ByVal 0&)
    Shift = (Number - 1) - curLine
    Call SendMessage(rtf.hwnd, EM_LINESCROLL, 0&, ByVal Shift)
End Sub

 Function GetCurrentLine(RichTextBox As RichTextBox) As Long
    Dim CurrentLine As Long
    Const EM_LINEFROMCHAR = &HC9
    CurrentLine = SendMessage(RichTextBox.hwnd, EM_LINEFROMCHAR, -1, 0&)
    GetCurrentLine = CurrentLine
End Function

Public Function GetTotalLines(RichTextBox As RichTextBox) As Long
    Const EM_GETLINECOUNT = &HBA
    GetTotalLines = SendMessage(RichTextBox.hwnd, EM_GETLINECOUNT, 0, 0&)
End Function

Public Function GetLineText(rtf As RichTextBox, line_index As Long) As String
    Dim lnglength As Long, linestart As Long
    Dim strbuffer As String
    Const EM_GETLINE = &HC4
    Const EM_LINELENGTH = &HC1
    Const EM_LINEINDEX = &HBB
    
    linestart = SendMessage(rtf.hwnd, EM_LINEINDEX, line_index, 0&)
    lnglength = SendMessage(rtf.hwnd, EM_LINELENGTH, linestart, 0)
    strbuffer = Space(lnglength)
    Call SendMessage(rtf.hwnd, EM_GETLINE, line_index, ByVal strbuffer)
    
    GetLineText = strbuffer
    
End Function

Private Sub List1_Click()
    On Error Resume Next
    
    Dim tmp As String
    Dim line As Long
    Dim index As Long
    
    index = ListSelIndex(List1)
    
    If index >= 0 Then
        tmp = List1.List(index)
        If InStr(1, tmp, ":") > 0 Then
            line = CLng(Split(tmp, ":")(0))
            ScrollToLine active_object, line
        End If
    End If
    
End Sub

Private Function ListSelIndex(lst As ListBox) As Long
    
    On Error GoTo hell
    
    For i = 0 To List1.ListCount
        If List1.Selected(i) Then
            ListSelIndex = i
            Exit Function
        End If
    Next

hell:
    ListSelIndex = -1
    
End Function

