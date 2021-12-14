Attribute VB_Name = "Registry"
Option Explicit

Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4

Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003

Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259

Global Const KEY_ALL_ACCESS = &H3F

Global Const REG_OPTION_NON_VOLATILE = 0

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String)
Private Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String)

Public Function DeleteKey(lPredefinedKey As Long, sKeyName As String)

    Dim lRetVal As Long
    Dim hKey As Long
    
   
    
    
    lRetVal = RegDeleteKey(lPredefinedKey, sKeyName)
    
End Function

Public Function DeleteValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)


       Dim lRetVal As Long
       Dim hKey As Long

       

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = RegDeleteValue(hKey, sValueName)
       RegCloseKey (hKey)
End Function

Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String

    Select Case lType
        Case REG_SZ
            sValue = vValue
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
        End Select

End Function





Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lRC As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError




    lRC = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lRC <> ERROR_NONE Then Error 5

    Select Case lType
     
        Case REG_SZ:
            sValue = String(cch, 0)
            lRC = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lRC = ERROR_NONE Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If

        Case REG_DWORD:
            lRC = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lRC = ERROR_NONE Then vValue = lValue
        Case Else
           
            lRC = -1
    End Select

QueryValueExExit:

    QueryValueEx = lRC
    Exit Function

QueryValueExError:

    Resume QueryValueExExit

End Function
Public Function CreateNewKey(lPredefinedKey As Long, sNewKeyName As String)

    
    
    Dim hNewKey As Long
    Dim lRetVal As Long
    
    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
End Function


Sub Main()
    'Examples of each function:
    'CreateNewKey HKEY_CURRENT_USER, "TestKey\SubKey1\SubKey2"
    'SetKeyValue HKEY_CURRENT_USER, "TestKey\SubKey1", "Test", "Testing, Testing", REG_SZ
    'MsgBox QueryValue(HKEY_CURRENT_USER, "TestKey\SubKey1", "Test")
    'DeleteKey HKEY_CURRENT_USER, "TestKey\SubKey1\SubKey2"
    'DeleteValue HKEY_CURRENT_USER, "TestKey\SubKey1", "Test"
End Sub


Public Function SetKeyValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)


       Dim lRetVal As Long
       Dim hKey As Long



       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
       RegCloseKey (hKey)

End Function

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)


       Dim lRetVal As Long
       Dim hKey As Long
       Dim vValue As Variant


       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = QueryValueEx(hKey, sValueName, vValue)
   
       QueryValue = vValue
       RegCloseKey (hKey)
End Function








