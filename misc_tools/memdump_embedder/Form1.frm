VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Embed memory dump into 32bit dll husk for quick disassembly at proper VA"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   5985
      TabIndex        =   4
      Top             =   180
      Width           =   1365
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   810
      TabIndex        =   2
      Top             =   675
      Width           =   8340
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   420
      Left            =   7605
      TabIndex        =   1
      Top             =   135
      Width           =   1590
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   810
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Text            =   "drag and drop file ex: [pid]_[hexbase].mem to auto parse base"
      Top             =   180
      Width           =   4470
   End
   Begin VB.Label Label2 
      Caption         =   "File"
      Height          =   285
      Left            =   90
      TabIndex        =   5
      Top             =   225
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "base"
      Height          =   285
      Left            =   5445
      TabIndex        =   3
      Top             =   225
      Width           =   465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim dummy As String
    
    On Error GoTo hell
    
    If Not FileExists(Text1) Then Exit Sub
    
    dummy = App.path & "\dummy.dll"
    If Not FileExists(dummy) Then
        MsgBox "Dummy dll not found?", vbCritical
        Exit Sub
    End If
    
    If Not isHex(Text2) Then
        MsgBox "Base address not valid hex string", vbCritical
        Exit Sub
    End If
    
    List1.Clear
    
    Dim b() As Byte
    Dim f As Long
    
    f = FreeFile
    Open Text1 For Binary As f
    ReDim b(LOF(f) + 1)
    Get f, , b()
    Close f
    
    Dim outFile As String
    outFile = GetParentFolder(Text1) & "\" & Text2 & "_dmp.dll"
    List1.AddItem "Generating " & outFile
    
    If FileExists(outFile) Then
        List1.AddItem "Deleting old file.."
        Kill outFile
    End If
    
    List1.AddItem "Copying dummy.dll to new file.."
    FileCopy dummy, outFile
    
    List1.AddItem "Appending memory dump.."
    f = FreeFile
    Open outFile For Binary As f
    Put f, LOF(f) + 1, b() 'append the memory dump to dummy header..
    
    While LOF(f) Mod &H10 <> 0 'file and section alignment has been set to 0x10..
        Put f, , CByte(0)
    Wend
    
    'imgsz = A....140
    'ep = D.......118
    'imgbase = C..124
    'vsz/rsz = B..1f0/1f8
    
    List1.AddItem "Fixing up PE attributes for easy disassembly.."
    Put f, &H140 + 1, CLng(LOF(f)) + 1 'image size
    Put f, &H124 + 1, CLng("&h" & Text2)  'image base
    Put f, &H1F0 + 1, CLng(UBound(b)) + 1 'virtual size
    Put f, &H1F8 + 1, CLng(UBound(b)) + 1 'raw size
    
    Close f
    
    List1.AddItem "Complete!"
    List1.AddItem ""
    List1.AddItem "You can manually set entry point it is 0 for now.."
    
    Exit Sub
    
hell:
    List1.AddItem "Error! " & Err.Description
    
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Text1 = Data.Files(1)
    a = InStrRev(Text1, "_")
    If a > 0 Then
        Text2 = Mid(Text1, a + 1)
        Text2 = Replace(Text2, ".mem", "")
        If Not isHex(Text2) Then Text2 = Empty
    End If
End Sub

Function isHex(x) As Boolean
    On Error GoTo hell
    Dim l As Long
    l = Hex(CLng("&h" & x))
    isHex = True
    Exit Function
hell:
End Function



Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function



Function GetParentFolder(path) As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
End Function

