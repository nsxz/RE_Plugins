VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4470
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   900
      Width           =   8655
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   360
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   135
      Width           =   5820
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6390
      TabIndex        =   0
      Top             =   135
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Const LANG_US = &H409

Private Sub Command1_Click()

    f = Text1
    If Not FileExists(f) Then Exit Sub
    Dim isBinary As Boolean
    Dim firstIndex As Long
    Dim ret() As String
    
    'f = "C:\Documents and Settings\david\Desktop\analysis\1812_2570000.mem"
    'f = "c:\noname.dat"
    'f = "c:\2.dat"
    
    dat = ReadFile(f)
    
    Dim fh As Long
    Dim b() As Byte
    Dim inc As Long
    Dim tmp As String
    Dim bb As Byte
    
    fh = FreeFile
    Open f For Binary As fh
    ReDim b(LOF(fh))
    Get fh, , b()
    Close fh
    
'seg000:00001C83 C6 45 5C 77                       mov     byte ptr [ebp+5Ch], 77h ; 'w'
'seg000:00001C87 C6 45 5D 61                       mov     byte ptr [ebp+5Dh], 61h ; 'a'
'seg000:00001C8B C6 45 5E 74                       mov     byte ptr [ebp+5Eh], 74h ; 't'
'
'seg000:00006F79 C6 85 C0 FE FF FF 43                          mov     [ebp+68h+var_1A8], 43h ; 'C'
'seg000:00006F80 C6 85 C1 FE FF FF 6F                          mov     [ebp+68h+var_1A7], 6Fh ; 'o'
'seg000:00006F87 C6 85 C2 FE FF FF 75                          mov     [ebp+68h+var_1A6], 75h ; 'u'
'seg000:00006F8E C6 85 C3 FE FF FF 6E                          mov     [ebp+68h+var_1A5], 6Eh ; 'n'

    Me.Caption = "1"
    push ret(), "method 1"
    a = InStr(dat, Chr(&HC6) & Chr(&H45))
    Do While a > 0
            
        bb = b(a + 2)
        
        If bb = 0 Then
            tmp = tmp & " "
        ElseIf isAscii(bb) Then
            tmp = tmp & Chr(bb)
        Else
            isBinary = True
            tmp = tmp & Chr(bb)
        End If
        
        If b(a + 3) <> &HC6 Then 'end of sequence..
            If Len(Trim(tmp)) > 0 Then
                If isBinary Then
                    If Len(tmp) > 4 Then push ret(), pad(a) & BinaryString(tmp)
                Else
                    push ret(), pad(a) & tmp
                End If
            End If
            tmp = Empty
            isBinary = False
        End If
        
        a = InStr(a + 1, dat, Chr(&HC6) & Chr(&H45))
               
        
    Loop
    
    push ret(), vbCrLf & vbCrLf & "method 2"
    Me.Caption = "2"
    isBinary = False
    a = InStr(dat, Chr(&HC6) & Chr(&H85))
    Do While a > 0
    
        bb = b(a + 5)
        
        If bb = 0 Then
             tmp = tmp & " "
        ElseIf isAscii(bb) Then
            tmp = tmp & Chr(bb)
        Else
            isBinary = True
            tmp = tmp & Chr(bb)
        End If
        
        If b(a + 6) <> &HC6 Then 'end of sequence..
            If Len(Trim(tmp)) > 0 Then
                If isBinary Then
                    If Len(tmp) > 4 Then push ret(), pad(a) & BinaryString(tmp)
                Else
                    push ret(), pad(a) & tmp
                End If
            End If
            tmp = Empty
            isBinary = False
        End If
        
        a = InStr(a + 1, dat, Chr(&HC6) & Chr(&H85))
               
        
    Loop
    
    'method3 cryptolocker 3
    's2:0040BC5E B8 5C 00 00 00                          mov     eax, 5Ch
    's2:0040BC63 66 89 45 B4                             mov     [ebp+var_4C], ax
    's2:0040BC67 B9 52 00 00 00                          mov     ecx, 52h
    's2:0040BC6C 66 89 4D B6                             mov     [ebp+var_4A], cx
    's2:0040BC70 BA 65 00 00 00                          mov     edx, 65h
    's2:0040BC75 66 89 55 B8                             mov     [ebp+var_48], dx
    's2:0040BC79 B8 67 00 00 00                          mov     eax, 67h
    's2:0040BC7E 66 89 45 BA                             mov     [ebp+var_46], ax
    's2:0040BC82 B9 69 00 00 00                          mov     ecx, 69h
    's2:0040BC87 66 89 4D BC                             mov     [ebp+var_44], cx
    's2:0040BC8B BA 73 00 00 00                          mov     edx, 73h
    's2:0040BC90 66 89 55 BE                             mov     [ebp+var_42], dx
    's2:0040BC94 B8 74 00 00 00                          mov     eax, 74h
    's2:0040BC99 66 89 45 C0                             mov     [ebp+var_40], ax
    's2:0040BC9D B9 72 00 00 00                          mov     ecx, 72h
    's2:0040BCA2 66 89 4D C2                             mov     [ebp+var_3E], cx
    's2:0040BCA6 BA 79 00 00 00                          mov     edx, 79h
    's2:0040BCAB 66 89 55 C4                             mov     [ebp+var_3C], dx
    's2:0040BCAF B8 5C 00 00 00                          mov     eax, 5Ch
    's2:0040BCB4 66 89 45 C6                             mov     [ebp+var_3A], ax
    's2:0040BCB8 33 C9                                   xor     ecx, ecx
    's2:0040BCBA 66 89 4D C8                             mov     [ebp+var_38], cx
    's2:0040BCBE BA 4D 00 00 00                          mov     edx, 4Dh


    
    
    Text2 = Join(ret, vbCrLf)
    Me.Caption = "done"
    
End Sub

Function pad(v, Optional leng = 8)
    On Error GoTo hell
    Dim x As String
    x = Hex(v)
    While Len(x) < leng
        x = "0" & x
    Wend
    pad = x & "  "
    Exit Function
hell:
    pad = x & "  "
End Function

Function isAscii(x As Byte) As Boolean
    If x >= 9 And x <= Asc("z") Then isAscii = True
End Function

Function BinaryString(str As String) As String
    Dim i As Long
    Dim ret As String
    Dim b() As Byte
    
    b() = StrConv(str, vbFromUnicode, LANG_US)
    
    For i = 0 To UBound(b)
         If b(i) < &H10 Then
            ret = ret & "0" & Hex(b(i))
         Else
             ret = ret & Hex(b(i))
         End If
    Next
    
    BinaryString = "Binary String (" & UBound(b) + 1 & " bytes): " & ret
    
End Function


Function FileExists(path) As Boolean
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Function ReadFile(filename)
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Text1 = Data.Files(1)
End Sub
