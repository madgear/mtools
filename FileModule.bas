Attribute VB_Name = "FileModule"
Option Explicit


Declare Function MoveWindow Lib "user32" _
                       (ByVal hWnd As Long, _
                        ByVal X As Long, ByVal Y As Long, _
                        ByVal nWidth As Long, ByVal nHeight As Long, _
                        ByVal bRepaint As Long) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                        (ByVal hWnd As Long, ByVal wMsg As Long, _
                         ByVal wParam As Long, lParam As Any) As Long
Public Const LB_INITSTORAGE = &H1A8
Public Const LB_ADDSTRING = &H180

Public Const WM_SETREDRAW = &HB
Public Const WM_VSCROLL = &H115
Public Const SB_BOTTOM = 7
Declare Function GetLogicalDrives Lib "kernel32" () As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
                        (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Public Const INVALID_HANDLE_VALUE = -1


Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
                        (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Const MaxLFNPath = 260

Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MaxLFNPath
        cShortFileName As String * 14
End Type



Public Function FileExist(strfilename As String) As Boolean
    On Error Resume Next
    FileExist = True
        If FileLen(strfilename) = 0 Then
            FileExist = False
        End If
End Function


