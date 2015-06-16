Attribute VB_Name = "Module1"
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
 
Global dlg As New clsCmnDlg
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long



Sub Main()
    Dim h As Long
    h = LoadLibrary(App.path & "\SciLexer.dll") 'override the default one with ours which has the asasm lexer
    frmRabcd.Show
End Sub

'this is a quick and dirty scan as a cheap shot..just to check basics..
Function cveScan(fPath As String) As String

    Dim cves() As String
    Dim matched As Boolean
    Dim ret() As String
    
    push cves, "CVE-2015-0310:new RegExp"
    push cves, "CVE-2015-0311:domainMemory,uncompress"
    push cves, "CVE-2015-0556:copyPixelsToByteArray"
    push cves, "CVE-2015-0313:createMessageChannel,createWorker"
    push cves, "CVE-2014-9163:parseFloat"
    push cves, "CVE-2014-0515:byteCode,Shader"
    push cves, "CVE-2014-0502:setSharedProperty,createWorker,.start,SharedObject"
    push cves, "CVE-2014-0497:writeUTFBytes,domainMemory"
    
    If Not FileExists(fPath) Then Exit Function
    
    dat = ReadFile(fPath)
    For Each cve In cves
        c = Split(cve, ":")
        checks = Split(c(1), ",")
        matched = False
        For Each k In checks
            If InStr(1, dat, k, vbTextCompare) > 0 Then matched = True Else matched = False
        Next
        If matched Then push ret, """" & cve & """    FILE: """ & FileNameFromPath(fPath) & """"
    Next
    
    cveScan = Join(ret, vbCrLf)
    
End Function


Function FileExists(path As String) As Boolean
  On Error GoTo hell
  Dim tmp As String
  tmp = Replace(path, "'", Empty)
  tmp = Replace(tmp, """", Empty)
  If Len(tmp) = 0 Then Exit Function
  If Dir(tmp, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  Exit Function
hell: FileExists = False
End Function

Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
init: ReDim ary(0): ary(0) = Value
End Sub

Function FolderExists(path As String) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbDirectory) <> "" Then FolderExists = True
End Function

Function GetParentFolder(path) As String
    If Len(path) = 0 Then Exit Function
    Dim tmp() As String
    Dim ub As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
End Function

Function GetBaseName(path As String) As String
    Dim tmp() As String
    Dim ub As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetBaseName = Mid(ub, 1, InStrRev(ub, ".") - 1)
    Else
       GetBaseName = ub
    End If
End Function



Function FileNameFromPath(fullpath) As String
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function

'Function GetFolderFiles(folder As String, Optional filter = "*.*", Optional retFullPath As Boolean = True) As String()
'   Dim fnames() As String
'
'   If Not FolderExists(folder) Then
'        'returns empty array if fails
'        GetFolderFiles = fnames()
'        Exit Function
'   End If
'
'   folder = IIf(Right(folder, 1) = "\", folder, folder & "\")
'   'If Left(filter, 1) = "*" Then extension = Mid(filter, 2, Len(filter))
'   'If Left(filter, 1) <> "." Then filter = "." & filter
'
'   fs = Dir(folder & filter, vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
'   While fs <> ""
'     If fs <> "" Then push fnames(), IIf(retFullPath = True, folder & fs, fs)
'     fs = Dir()
'   Wend
'
'   GetFolderFiles = fnames()
'End Function

Function GetFolderFiles(folderPath As String, Optional filter As String = "*", Optional retFullPath As Boolean = True, Optional recursive As Boolean = False) As String()
   Dim fnames() As String
   Dim fs As String
   Dim folders() As String
   Dim i As Integer
   
   If Not FolderExists(folderPath) Then
        'returns empty array if fails
        GetFolderFiles = fnames()
        Exit Function
   End If
   
   folderPath = IIf(Right(folderPath, 1) = "\", folderPath, folderPath & "\")
   
   fs = Dir(folderPath & filter, vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
   While fs <> ""
     If fs <> "" Then push fnames(), IIf(retFullPath = True, folderPath & fs, fs)
     fs = Dir()
   Wend
   
   If recursive Then
        folders() = GetSubFolders(folderPath)
        If Not AryIsEmpty(folders) Then
            For i = 0 To UBound(folders)
                FolderEngine folders(i), fnames(), filter
            Next
        End If
        If Not retFullPath Then
            For i = 0 To UBound(fnames)
                fnames(i) = Replace(fnames(i), folderPath, Empty) 'make relative path from base
            Next
        End If
    End If
   
   GetFolderFiles = fnames()
End Function

Private Sub FolderEngine(fldrpath As String, ary() As String, Optional filter As String = "*")

    Dim files() As String
    Dim folders() As String
    Dim i As Long
     
    files = GetFolderFiles(fldrpath, filter)
    folders = GetSubFolders(fldrpath)
        
    If Not AryIsEmpty(files) Then
        For i = 0 To UBound(files)
            push ary, files(i)
        Next
    End If
    
    If Not AryIsEmpty(folders) Then
        For i = 0 To UBound(folders)
             FolderEngine folders(i), ary, filter
        Next
    End If
    
End Sub

Function DeleteFolder(folderPath As String, Optional force As Boolean = True) As Boolean
 On Error GoTo failed
   Call delTree(folderPath, force)
   RmDir folderPath
   DeleteFolder = True
 Exit Function
failed:  DeleteFolder = False
End Function

Private Sub delTree(folderPath As String, Optional force As Boolean = True)
   Dim sfi() As String, sfo() As String, i As Integer
   sfi() = GetFolderFiles(folderPath)
   sfo() = GetSubFolders(folderPath)
   If Not AryIsEmpty(sfi) And force = True Then
        For i = 0 To UBound(sfi)
            DeleteFile sfi(i)
        Next
   End If
   
   If Not AryIsEmpty(sfo) And force = True Then
        For i = 0 To UBound(sfo)
            Call DeleteFolder(sfo(i), True)
        Next
   End If
End Sub

Function DeleteFile(fPath As String) As Boolean
 On Error GoTo hadErr
    
    Dim attributes As VbFileAttribute

    attributes = GetAttr(fPath)
    If (attributes And vbReadOnly) Then
        attributes = attributes - vbReadOnly
        SetAttr fPath, attributes
    End If

    Kill fPath
    DeleteFile = True
    
 Exit Function
hadErr:
'MsgBox "DeleteFile Failed" & vbCrLf & vbCrLf & fpath
DeleteFile = False
End Function

Sub WriteFile(path As String, it As Variant)
    Dim f As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub



Function GetSubFolders(folder As String, Optional retFullPath As Boolean = True) As String()
    Dim fnames() As String
    
    If Not FolderExists(folder) Then
        'returns empty array if fails
        GetSubFolders = fnames()
        Exit Function
    End If
    
   If Right(folder, 1) <> "\" Then folder = folder & "\"

   fd = Dir(folder, vbDirectory)
   While fd <> ""
     If Left(fd, 1) <> "." Then
        If (GetAttr(folder & fd) And vbDirectory) = vbDirectory Then
           push fnames(), IIf(retFullPath = True, folder & fd, fd)
        End If
     End If
     fd = Dir()
   Wend
   
   GetSubFolders = fnames()
End Function


Function ReadFile(filename) As Variant
  Dim f As Long
  Dim temp As Variant
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Function RandomNum() As Long
    Dim tmp As Long
    Dim tries As Long
    
    On Error GoTo again
tryit:

    Randomize
    tmp = Round(Timer * Now * Rnd(), 0)
    RandomNum = tmp
    
    Exit Function
again:
    
    If tries < 10 Then
        tries = tries + 1
        GoTo tryit
    End If
    
End Function

Function GetFreeFileName(ByVal folder As String, Optional extension = ".txt") As String
    
    On Error GoTo handler 'can have overflow err once in awhile :(
    Dim i As Integer
    Dim tmp As String

    If Not FolderExists(folder) Then Exit Function
    If Right(folder, 1) <> "\" Then folder = folder & "\"
    If Left(extension, 1) <> "." Then extension = "." & extension
    
again:
    Do
      tmp = folder & RandomNum() & extension
    Loop Until Not FileExists(tmp)
    
    GetFreeFileName = tmp
    
Exit Function
handler:

    If i < 10 Then
        i = i + 1
        GoTo again
    End If
    
End Function

Public Function GetShortName(sFile As String) As String
    Dim sShortFile As String * 300
    Dim lResult As Long

    If Not FileExists(sFile) Then
        MsgBox "GetshortName file must exist to work..: " & sFile
        GetShortName = sFile
        Exit Function
    End If
    
    'file must exist or this will fail...
    lResult = GetShortPathName(sFile, sShortFile, Len(sShortFile))

    'Trim out unused characters from the string.
    GetShortName = Left$(sShortFile, lResult)

    If Not FileExists(GetShortName) Then GetShortName = sFile
End Function

Sub SetLiColor(li As ListItem, newcolor As Long)
    Dim f As ListSubItem
'    On Error Resume Next
    li.ForeColor = newcolor
    For Each f In li.ListSubItems
        f.ForeColor = newcolor
    Next
End Sub

Sub LV_LastColumnResize(lv As ListView)
    On Error Resume Next
    lv.ColumnHeaders(lv.ColumnHeaders.Count).Width = lv.Width - lv.ColumnHeaders(lv.ColumnHeaders.Count).Left - 100
End Sub

Public Sub LV_ColumnSort(ListViewControl As ListView, Column As ColumnHeader)
     On Error Resume Next
    With ListViewControl
       If .SortKey <> Column.Index - 1 Then
             .SortKey = Column.Index - 1
             .SortOrder = lvwAscending
       Else
             If .SortOrder = lvwAscending Then
              .SortOrder = lvwDescending
             Else
              .SortOrder = lvwAscending
             End If
       End If
       .Sorted = -1
    End With
End Sub

Function pad(v, Optional l As Long = 4)
    On Error GoTo hell
    Dim x As Long
    x = Len(v)
    If x < l Then
        pad = String(l - x, " ") & v
    Else
hell:
        pad = v
    End If
End Function

Function isDecimalNumber(x) As Boolean
    
    'Debug.Print isDecimalNumber("32")    'true
    'Debug.Print isDecimalNumber("32 ")   'true
    'Debug.Print isDecimalNumber("232a ") 'false
    ' Stop
     
    On Error GoTo hell
    Dim l As Long
    
    For i = 1 To Len(x) - 1
        c = Mid(x, i, 1)
        If Not IsNumeric(c) Then Exit Function
    Next
    
    l = CLng(x)
    isDecimalNumber = True
    
hell:
    Exit Function
    
End Function

Function StringOpcodesToBytes(OpCodes) As Byte()
    
    'Debug.Print StrConv(StringOpcodesToBytes("41 42 43 44"), vbUnicode)
    'Stop
    
    On Error Resume Next
    Dim b() As Byte
    
    tmp = Split(Trim(OpCodes), " ")
    ReDim b(UBound(tmp))
    
    For i = 0 To UBound(tmp)
        b(i) = CByte(CInt("&h" & tmp(i)))
    Next
    
    StringOpcodesToBytes = b()
    
End Function

Function lpad(x, Optional sz = 8)
    a = Len(x) - sz
    If a < 0 Then
        lpad = x & Space(Abs(a))
    Else
        lpad = x
    End If
End Function

Function objKeyExistsInCollection(c As Collection, k As String) As Boolean
    On Error GoTo hell
    Set x = c(k)
    objKeyExistsInCollection = True
hell:
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim x As Long
    x = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function
 
