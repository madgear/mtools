Attribute VB_Name = "ProcessModule"
Sub ForceKill(pname As String)

If CheckOS = "WinNT" Then

Dim processes As New ProcessList
For i = 0 To processes.processCount - 1
If UCase(processes.ProcessName(i)) = UCase(pname) Then
processes.KillProcess processes.ProcessHandle(i)
End If
Next
Set processes = Nothing

Else

With mainform

.Timer1.Enabled = False
.KillApp ("none")

For i = 0 To .memlist.ListCount - 1
If UCase(.memlist.List(i)) = UCase(pname) Then
.KillApp pname
End If
Next

.KillApp ("none")
.Timer1.Enabled = True

End With

End If
End Sub

Function ProcessExist(pname) As Boolean
Dim processes As New ProcessList
Dim dummy1
dummy1 = 0
For i = 0 To processes.processCount - 1
If UCase(processes.ProcessName(i)) = UCase(pname) Then
dummy1 = 1
End If
Next
Set processes = Nothing
If dummy1 = 1 Then
ProcessExist = True
Else
ProcessExist = False
End If
End Function


