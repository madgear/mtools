Attribute VB_Name = "inifile"
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Global optfile As String
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
