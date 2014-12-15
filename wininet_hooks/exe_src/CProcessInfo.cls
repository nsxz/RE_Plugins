VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProcessInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:    David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA

'Used in several projects do not change interface!

Private Declare Function NtQuerySystemInformation Lib "ntdll.dll" ( _
    ByVal dwInfoType As Long, _
    ByRef lpStructure As Any, _
    ByVal dwSize As Long, _
    dwReserved As Long) As Long

Private Type SYSTEM_MODULE_INFORMATION
    reserved(1) As Long                'ULONG reserved[2];
    base As Long                       'PVOID Base;
    size As Long                       'ULONG Size;
    flags As Long                      'ULONG Flags;
    index As Integer                   'USHORT Index;
    unkn As Integer                    'USHORT Unknown;
    lcount As Integer                  'USHORT LoadCount;
    modoffset As Integer               'USHORT ModuleNameOffset;
    ImageName As String * 256          'CHAR ImageName[256];
End Type

Private Const SYSMODINFO_SPECIFIER = 11

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Enum ProcessAccessTypes
  PROCESS_TERMINATE = (&H1)
  PROCESS_CREATE_THREAD = (&H2)
  PROCESS_SET_SESSIONID = (&H4)
  PROCESS_VM_OPERATION = (&H8)
  PROCESS_VM_READ = (&H10)
  PROCESS_VM_WRITE = (&H20)
  PROCESS_DUP_HANDLE = (&H40)
  PROCESS_CREATE_PROCESS = (&H80)
  PROCESS_SET_QUOTA = (&H100)
  PROCESS_SET_INFORMATION = (&H200)
  PROCESS_QUERY_INFORMATION = (&H400)
'  STANDARD_RIGHTS_REQUIRED = &HF0000
  SYNCHRONIZE = &H100000
  PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
End Enum

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const ALL_ACCESS = &H1F0FFF
Private Const TH32CS_SNAPHEAPLIST = 1
Private Const TH32CS_SNAPPROCESS = 2
Private Const TH32CS_SNAPTHREAD = 4
Private Const TH32CS_SNAPMODULE = 8
Private Const TOKEN_QUERY = 8&
Private Const TOKEN_USER = 1
Private Const SPARE_LEN = 512
Private Const MEM_RELEASE = &H8000
Private Const MEM_COMMIT = &H1000
Private Const MEM_RESERVE = &H2000
Private Const PAGE_EXECUTE_READWRITE = &H40
Private Const PAGE_EXECUTE_READ = &H20
Private Const PAGE_EXECUTE_WRITECOPY = &H80
Private Const PAGE_GUARD = &H100
Private Const PAGE_NOACCESS = &H1
Private Const PAGE_NOCACHE = &H200
Private Const PAGE_READONLY = &H2
Private Const PAGE_READWRITE = &H4
Private Const PAGE_WRITECOMBINE = &H400
Private Const PAGE_WRITECOPY = &H8
Private Const MEM_MAPPED = &H40000
Private Const MEM_IMAGE = &H1000000
Private Const MEM_PRIVATE = &H20000
Private Const MEM_PHYSICAL = &H400000
'Private Const PROCESS_TERMINATE = 1

Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Byte, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
                           
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleInformation Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByRef lpmodinfo As LPMODULEINFO, ByVal cb As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long

Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal Sid As Long, ByVal name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
Private Declare Function VirtualQueryEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function GetLastError Lib "kernel32.dll" () As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As TOKEN_PRIVILEGES, ByRef ReturnLength As Long) As Long

Private Type MEMORY_BASIC_INFORMATION
    BaseAddress As Long
    AllocationBase As Long
    AllocationProtect As Long
    RegionSize As Long
    State As Long
    Protect As Long
    lType As Long
End Type

Private Type SID_AND_ATTRIBUTES
       Sid As Long
       Attributes As Long
       Spare(SPARE_LEN) As Byte
End Type

Private Type LPMODULEINFO
    lpBaseOfDll As Long
    SizeOfImage As Long
    EntryPoint As Long
End Type

Private Type PROCESS_BASIC_INFORMATION
    ExitStatus As Long
    PebBaseAddress As Long
    AffinityMask As Long
    BasePriority As Long
    UniqueProcessId As Long
    InheritedFromUniqueProcessId As Long
End Type

Private Type strucPEB
    filler(15) As Byte
    InfoBlockAdderss As Long
End Type

Private Type INFOBLOCK_
    filler(67) As Byte
    wszCmdLineAddress As Long
End Type

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type

Function GetRunningProcesses() As Collection 'of CProcess classes
    
    Dim m_col As New Collection
    Dim myProcess As PROCESSENTRY32
    Dim mySnapshot As Long
    Dim proc As CProcess
    
    myProcess.dwSize = Len(myProcess)
    mySnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    
    ProcessFirst mySnapshot, myProcess
    
    Set proc = New CProcess
    With proc
        '.path = GetProcessPath(myProcess.th32ProcessID)
        .pid = myProcess.th32ProcessID
        .ParentPID = myProcess.th32ParentProcessID
        '.CmdLine = GetProcessCmdLine(.pid)
        '.User = GetProcessUser(.pid)
    End With
    
    m_col.Add proc, "pid:" & myProcess.th32ProcessID
    
     
    While ProcessNext(mySnapshot, myProcess)
        Set proc = New CProcess
        With proc
            .path = myProcess.szexeFile
            .pid = myProcess.th32ProcessID
            .ParentPID = myProcess.th32ParentProcessID
            '.CmdLine = GetProcessCmdLine(.pid)
            '.User = GetProcessUser(.pid)
        End With
        
        m_col.Add proc, "pid:" & myProcess.th32ProcessID
    Wend

    Set GetRunningProcesses = m_col
    
End Function

Function GetSeDebug() As Boolean
    Dim hToken As Long, hProcess As Long, lret As Long
    Dim tkp As TOKEN_PRIVILEGES
    
    Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
    Const TOKEN_QUERY As Long = &H8

    hProcess = GetCurrentProcess()
    OpenProcessToken hProcess, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken
    LookupPrivilegeValue "", "SeDebugPrivilege", tkp.TheLuid

    tkp.PrivilegeCount = 1
    tkp.Attributes = 2 'SE_PRIVILEGE_ENABLED
    
    AdjustTokenPrivileges hToken, False, tkp, Len(tkp), tkp, lret
        
    If GetLastError = 0 Then GetSeDebug = True

End Function

Private Sub Class_Initialize()
    GetSeDebug
End Sub
