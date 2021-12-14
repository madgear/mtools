VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type LUID: lowpart As Long: highpart As Long: End Type
Private Type TOKEN_PRIVILEGES: PrivilegeCount As Long: LuidUDT As LUID: Attributes As Long: End Type

Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Const PROCESS_TERMINATE As Long = &H1
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal moduleName As String, ByVal nSize As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16


Private Type typProcess
    Handle As Long
    exeName As String
End Type
Private Process() As typProcess, intProcessCount As Integer
Event FoundProcess(Index As Integer, Handle As Long, exeName As String)


Public Function processCount() As Integer
 processCount = intProcessCount
End Function


Public Property Get ProcessName(Index) As String
On Error GoTo bailout
 ProcessName = Process(Index).exeName
bailout:
If Err.Number <> 0 Then Err.Raise 1, "ProcessName", "invalid index, index [" & Index & "] not found"
End Property


Public Property Get ProcessHandle(Index) As Long
On Error GoTo bailout
 ProcessHandle = Process(Index).Handle
bailout:
If Err.Number <> 0 Then Err.Raise 1, "ProcessHandle", "invalid index, index [" & Index & "] not found"
End Property


Public Function CheckProcesses() As Integer

CheckWindowsVersion

Dim cb As Long
Dim cbNeeded As Long
Dim ProcessIDs() As Long
Dim lngResult As Long
Dim moduleName As String
Dim hProcess As Long
Dim i As Long
clearProcesses
cb = 32 * 4
Do
 ReDim ProcessIDs(cb / 4)
 lngResult = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
 If lngResult <> 1 Then Exit Function
 If cb = cbNeeded Then cb = cb * 2 Else Exit Do
Loop
For i = 1 To cbNeeded / 4
   DoEvents

   hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(i))

   If hProcess <> 0 And ProcessIDs(i) <> 0 Then

       moduleName = Space(256)
       lngResult = GetModuleFileNameExA(hProcess, 0, moduleName, Len(moduleName))

       If lngResult <> 0 Then
         moduleName = Left(moduleName, lngResult)

         Call appendProcess(ProcessIDs(i), moduleName)
       End If
   End If

   lngResult = CloseHandle(hProcess)
Next
End Function


Public Sub clearProcesses()
 ReDim Process(0) As typProcess
 intProcessCount = 0
End Sub


Private Function appendProcess(Handle As Long, exeName As String) As Integer
 ReDim Preserve Process(intProcessCount)
 With Process(intProcessCount)
    .exeName = exeName
    .Handle = Handle
 End With
 appendProcess = intProcessCount
 intProcessCount = intProcessCount + 1
 
 RaiseEvent FoundProcess(appendProcess, Handle, exeName)
End Function

Private Sub CheckWindowsVersion()
If GetVersion < 0 Then Err.Raise 2, "ProcessList", "This module can only be run on NT based operating systems"
End Sub


Public Function KillProcess(ByVal hProcessID As Long, Optional ByVal ExitCode As Long) As Boolean
    Dim hToken As Long
    Dim hProcess As Long
    Dim tp As TOKEN_PRIVILEGES
    
    If GetVersion() >= 0 Then
        If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) = 0 Then
            GoTo CleanUp
        End If
        
        If LookupPrivilegeValue("", "SeDebugPrivilege", tp.LuidUDT) = 0 Then
            GoTo CleanUp
        End If
    
        tp.PrivilegeCount = 1
        tp.Attributes = SE_PRIVILEGE_ENABLED
    
        If AdjustTokenPrivileges(hToken, False, tp, 0, ByVal 0&, ByVal 0&) = 0 Then
            GoTo CleanUp
        End If
    End If
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, hProcessID)
    If hProcess Then
        KillProcess2 = (TerminateProcess(hProcess, ExitCode) <> 0)

        CloseHandle hProcess
    End If
    
    If GetVersion() >= 0 Then

        tp.Attributes = 0
        AdjustTokenPrivileges hToken, False, tp, 0, ByVal 0&, ByVal 0&
        
CleanUp:
        If hToken Then CloseHandle hToken
    End If
End Function

Private Sub Class_Initialize()
 Me.CheckProcesses
End Sub

Private Sub Class_Terminate()
 clearProcesses
End Sub


