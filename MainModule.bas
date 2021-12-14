Attribute VB_Name = "MainModule"
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
'Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                     ByVal wMsg As Long, ByVal wParam As Long, ByVal sParam As String) As Long

Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const LB_SETCURSEL = &H186
Public Const LB_DELETESTRING = &H182
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETCURSEL = &H188
Public Const LB_INSERTSTRING = &H181

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Global DefFile As String
Global pTitle As String

Function WriteIni(Filename As String, Section As String, keyword As String, Value As String)
   Y = WritePrivateProfileString(Section, keyword, Value, Filename)
End Function
Public Function GetINIFile(ByVal fname As String, ByVal szSection As String, ByVal szField As String)
On Error GoTo Err_GetINIFile

    Dim strProcName As String
    Dim nRet As Integer
    Dim szFileName As String
    Dim szBuffer As String
    Dim szDefault As String
    Dim nTempLength As Integer

    strProcName = "GetINIFile"

    szDefault = ""
    szBuffer = String(80, " ")
    nRet = GetPrivateProfileString(szSection, szField, szDefault, szBuffer, Len(szBuffer), fname)
    If nRet > 0 Then
        GetINIFile = Left(szBuffer, nRet)
    Else
        GetINIFile = szDefault
    End If

Exit_GetINIFile:
    Exit Function

Err_GetINIFile:
    MsgBox Err.Number & Err.Description & vbCrLf & strProcName, vbCritical
    GoTo Exit_GetINIFile
    
End Function

Sub DragObject(ByVal hWnd As Long)
ReleaseCapture
SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Function CheckOS() As String
If GetVersion < 0 Then
CheckOS = "Win9x"
Else
CheckOS = "WinNT"
End If
End Function



Sub RemoveExcess(v1 As ListBox, v2 As ListBox, v3 As ListBox)

    Dim vStrPos As Variant
    Dim sPhrase As String
    Dim lRC As Long
   Dim processes As New ProcessList
DoEvents
  v3.Clear
  For no = 0 To v2.ListCount - 1
  
        sPhrase = v2.List(no)
        With v1
        vStrPos = SendMessageByString&(.hWnd, LB_FINDSTRINGEXACT, 0, sPhrase)
        If vStrPos = -1 Then
        DoEvents
            v3.AddItem sPhrase
        DoEvents
            lRC = SendMessage(.hWnd, LB_SETCURSEL, -1, 0)
        Else
            lRC = SendMessage(.hWnd, LB_SETCURSEL, vStrPos, 0)
        End If
        End With
    
  Next
DoEvents

End Sub
