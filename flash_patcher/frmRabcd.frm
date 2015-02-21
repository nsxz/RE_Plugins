VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRabcd 
   Caption         =   "Form2"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14655
   LinkTopic       =   "Form2"
   ScaleHeight     =   9465
   ScaleWidth      =   14655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdIntegrate 
      Caption         =   "integrate"
      Height          =   375
      Left            =   13365
      TabIndex        =   10
      Top             =   540
      Width           =   1140
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   10575
      TabIndex        =   9
      Top             =   540
      Width           =   1320
   End
   Begin VB.CommandButton cmdReasm 
      Caption         =   "reassemble"
      Height          =   375
      Left            =   12015
      TabIndex        =   8
      Top             =   540
      Width           =   1230
   End
   Begin VB.TextBox txtMod 
      Height          =   330
      Left            =   720
      TabIndex        =   7
      Top             =   450
      Width           =   7305
   End
   Begin VB.CommandButton cmdbrowse 
      Caption         =   "..."
      Height          =   330
      Left            =   8190
      TabIndex        =   5
      Top             =   45
      Width           =   735
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   8295
      Left            =   90
      TabIndex        =   4
      Top             =   900
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   14631
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   531
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdDissassemble 
      Caption         =   "Disasm"
      Height          =   330
      Left            =   9045
      TabIndex        =   3
      Top             =   45
      Width           =   1095
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   90
      Width           =   7305
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   8250
      Left            =   3690
      TabIndex        =   0
      Top             =   945
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   14552
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmRabcd.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "OutFile"
      Height          =   330
      Left            =   90
      TabIndex        =   6
      Top             =   495
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "File"
      Height          =   285
      Left            =   270
      TabIndex        =   1
      Top             =   90
      Width           =   420
   End
End
Attribute VB_Name = "frmRabcd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dp As String
Dim curFile As String
Dim curNode As Node
Dim isDirty As Boolean

Public Sub LoadFile(pth As String)
    txtFile = pth
    If FileExists(pth) Then cmdDissassemble_Click
    Me.Visible = True
End Sub

Private Sub cmdBrowse_Click()
    Dim x As String
    x = dlg.OpenDialog(AllFiles, , , Me.hWnd)
    If Len(x) = 0 Then Exit Sub
    txtFile = x
    cmdDissassemble_Click
End Sub

Private Sub cmdDissassemble_Click()
    a = "ExportABC: " & GetCommandOutput(dp & "\abcexport.exe """ & txtFile & """") & vbCrLf
    a = a & "swfbinexport: " & GetCommandOutput(dp & "\swfbinexport.exe """ & txtFile & """")
    rtf.Text = a
    
    Set curNode = Nothing
    tv.Nodes.Clear
    curFile = Empty
    
    Dim tmp() As String, pf As String, i As Long, ff As String, newf As String
    Dim bn As String
    
    bn = FileNameFromPath(txtFile)
    pf = GetParentFolder(txtFile)
    txtMod = pf & "\mod_" & bn
    
    tmp() = GetFolderFiles(pf, "*.abc")
    
    'now disassemble each abc file..
    For i = 0 To UBound(tmp)
        ff = tmp(i)
        'lv.ListItems.Add , , FileNameFromPath(ff)
        rtf.Text = rtf.Text & GetCommandOutput(dp & "\rabcdasm.exe """ & ff & """") & vbCrLf
        newf = pf & "\" & GetBaseName(FileNameFromPath(ff))
        If FolderExists(newf) Then addsubtree newf
        
    Next
    
    tmp() = GetFolderFiles(pf, "*.bin") 'should be [parent file name]*.bin
    Dim n As Node, n2 As Node
    
    If Not AryIsEmpty(tmp) Then
        Set n = tv.Nodes.Add(, , , "BinaryData")
        For i = 0 To UBound(tmp)
            ff = tmp(i)
            Set n2 = tv.Nodes.Add(n, tvwChild, , FileNameFromPath(ff))
        Next
    End If
    
    
    
End Sub

Private Sub cmdIntegrate_Click()
    MsgBox "todo:"
End Sub

Private Sub cmdReasm_Click()
    If FileExists(curFile) Then
        If isDirty Then cmdSave_Click
        MsgBox "todo: walk nodes up and find the -x.main.asasm file that contains this script for this abc block index.."
        't = GetCommandOutput(dp & "\rabcasm.exe """ & curFile & """", True, True)
        'MsgBox "command output: " & t
    End If
End Sub

Private Sub cmdSave_Click()
    If FileExists(curFile) Then
        WriteFile curFile, rtf.Text
        curNode.BackColor = vbBlue
        isDirty = False
    End If
End Sub

Private Sub Form_Load()
    dp = App.path & "\RABCDAsm_v1.17"
    If Not FolderExists(dp) Then
        cmdDissassemble.Enabled = False
        MsgBox "could not find RABCDAsm_v1.17 folder", vbInformation
    End If
End Sub

Sub addsubtree(pth As String, Optional pn As Node = Nothing)
    Dim n As Node, n2 As Node, ff() As String
    
    If pn Is Nothing Then
        Set n = tv.Nodes.Add(, , , FileNameFromPath(pth))
    Else
        Set n = tv.Nodes.Add(pn, tvwChild, , FileNameFromPath(pth))
    End If
    
    ff() = GetFolderFiles(pth)
    If Not AryIsEmpty(ff) Then
        For Each f In ff
            Set n2 = tv.Nodes.Add(n, tvwChild, , FileNameFromPath(f))
            n2.Tag = f
        Next
    End If
    
    ff() = GetSubFolders(pth)
    If Not AryIsEmpty(ff) Then
        For Each f In ff
            addsubtree CStr(f), n
        Next
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    rtf.Width = Me.Width - rtf.Left - 200
    rtf.Height = Me.Height - rtf.Top - 400
    tv.Height = rtf.Height
End Sub

Private Sub rtf_Change()
    isDirty = True
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim f As String
    
    Set curNode = Node
    curFile = Empty
    
    f = CStr(Node.Tag)
    
    If FileExists(f) Then
        rtf.Text = ReadFile(f)
        curFile = f
        isDirty = False
        rtf.SelStart = 1
    End If
    
End Sub
