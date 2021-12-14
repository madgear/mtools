Attribute VB_Name = "Remover"
Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Enum FileSysTyp
    File = 0
    Folder = 1
    Drive = 2
End Enum

Sub DestroyFile(sFileName As String)

On Error GoTo errhand
    Dim Block1 As String, Block2 As String, Blocks As Long
    Dim hFileHandle As Integer, iLoop As Long, offset As Long
    
    Const BLOCKSIZE = 4096
    Block1 = String(BLOCKSIZE, "X")
    Block2 = String(BLOCKSIZE, " ")
 
    hFileHandle = FreeFile
    Open sFileName For Binary As hFileHandle
    Blocks = (LOF(hFileHandle) \ BLOCKSIZE) + 1


    For iLoop = 1 To Blocks
        offset = Seek(hFileHandle)
        Put hFileHandle, , Block1
        Put hFileHandle, offset, Block2
    Next iLoop

    Close hFileHandle

    Kill sFileName
    
errhand:
If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, pTitle
End Sub

Sub SetAttr(sFileName As String)
On Error GoTo errhand:

If FileExist(sFileName) = True Then
SetFileAttributes sFileName, vbNormal
SetFileAttributes sFileName, vbArchive
End If
errhand:
If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, pTitle
End Sub

Sub RemoveFile(sname As String)
On Error GoTo errhand

ForceKill sname
SetAttr sname
DestroyFile sname

errhand:
If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, pTitle
End Sub


Sub VirusCheck(fname As String, vname As String, ByVal vsize)
On Error Resume Next
Dim tmpFile As String

tmpFile = fname & ".tmp"
If vsize = "FSIZE" Then
RemoveFile fname
Else
    If FileLen(fname) <= Val(vsize) Then
    RemoveFile fname
    Else
    
        Select Case UCase(vname)
    
        Case UCase("Win32.Spocls.1")
        CopyFile fname, tmpFile, Val(vsize), 28
        Case Else
        CopyFile fname, tmpFile, Val(vsize)
        End Select
    
    End If
End If

End Sub

Sub regfix()
On Error GoTo errhand

Dim regStr() As String

SetKeyValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "Explorer.exe", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Hidden", 1, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden", 0, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", 0, REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "Explorer.exe", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "load", "", REG_SZ
SetKeyValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "Homepage", 0, REG_DWORD

DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDriveTypeAutoRun"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"

SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "Explorer.exe", REG_SZ
SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Hidden", 1, REG_DWORD
SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden", 0, REG_DWORD
SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", 0, REG_DWORD

DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun"
DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDriveTypeAutoRun"
DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr"
DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD"
DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"
DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"

With mainform

For no = 0 To .reglist.ListCount - 1

regStr = Split(.reglist.List(no), ",")

    If regStr(2) = "HKCU" Then
     
      If regStr(3) = "0" Then
        
        DeleteValue HKEY_CURRENT_USER, regStr(0), regStr(1)
      
      Else
         
         Select Case regStr(4)
         Case "DWORD"
         SetKeyValue HKEY_CURRENT_USER, regStr(0), regStr(1), regStr(5), REG_DWORD
         Case "STRING"
         SetKeyValue HKEY_CURRENT_USER, regStr(0), regStr(1), regStr(5), REG_SZ
         End Select
                         
      End If
     
    ElseIf regStr(2) = "HKLM" Then


      If regStr(3) = "0" Then
       
        DeleteValue HKEY_LOCAL_MACHINE, regStr(0), regStr(1)
      
      Else
         
         Select Case regStr(4)
         Case "DWORD"
         SetKeyValue HKEY_LOCAL_MACHINE, regStr(0), regStr(1), regStr(5), REG_DWORD
         Case "STRING"
         SetKeyValue HKEY_LOCAL_MACHINE, regStr(0), regStr(1), regStr(5), REG_SZ
         End Select
                         
      End If


    ElseIf regStr(2) = "HCR" Then


      If regStr(3) = "0" Then
       
        DeleteValue HKEY_CLASSES_ROOT, regStr(0), regStr(1)
      
      Else
         
         Select Case regStr(4)
         Case "DWORD"
         SetKeyValue HKEY_CLASSES_ROOT, regStr(0), regStr(1), regStr(5), REG_DWORD
         Case "STRING"
         SetKeyValue HKEY_CLASSES_ROOT, regStr(0), regStr(1), regStr(5), REG_SZ
         End Select
                         
      End If
      
      
    ElseIf regStr(2) = "HU" Then


      If regStr(3) = "0" Then
       
        DeleteValue HKEY_USERS, regStr(0), regStr(1)
      
      Else
         
         Select Case regStr(4)
         Case "DWORD"
         SetKeyValue HKEY_USERS, regStr(0), regStr(1), regStr(5), REG_DWORD
         Case "STRING"
         SetKeyValue HKEY_USERS, regStr(0), regStr(1), regStr(5), REG_SZ
         End Select
                         
      End If

  End If

Next

End With

errhand:
If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, pTitle
End Sub

Sub CopyFile(ByVal F1, ByVal F2, vsize As Long, Optional dSize As Long)
'On Error Resume Next

Dim b() As Byte
Dim c() As Byte

Dim i As Long

ForceKill CStr(F1)
SetAttr CStr(F1)

  
  Open F1 For Binary Access Read As #1
  Open F2 For Binary Access Write As #2
  
  ReDim b(1 To LOF(1))
  ReDim c(1 To ((LOF(1) - vsize) - dSize))
  
  Get #1, 1, b
  
  For i = (vsize + 1) To (FileLen(F1) - dSize)
  DoEvents
  c(i - vsize) = b(i)
  DoEvents
  Next
  
  Close #1
  Erase b
  
  Put #2, 1, c
  Close #2
  Erase c
  
  RemoveFile CStr(F1)
  Copyf F2, F1
  RemoveFile CStr(F2)
  
End Sub

Sub Copyf(ByVal F12, ByVal F22)
On Error Resume Next
Dim b() As Byte
  Open F12 For Binary Access Read As #1
  Open F22 For Binary Access Write As #2
  ReDim b(1 To LOF(1))
  Get #1, 1, b
  Put #2, 1, b
  Close #1
  Close #2
  Erase b
End Sub
