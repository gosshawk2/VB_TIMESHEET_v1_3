Imports System.IO
Imports System.IO.File
Public Class ErrorLog
    Public Sub LogError(ByVal Logfilename As String, ByVal LogErrorPath As String, ByVal ErrorMessage As String, Optional ByVal RoutineName As String = "", Optional ByVal infoPrefix As String = "")
        Dim Log_Filename As String = LogErrorPath & "\Errors" & Logfilename
        Dim sw As StreamWriter = AppendText(Log_Filename)

        If Len(infoPrefix) > 0 Then
            sw.WriteLine(Now() & " " & infoPrefix & ", {" & RoutineName & "} ," & ErrorMessage)
        Else
            sw.WriteLine(Now() & " {" & RoutineName & "} ," & ErrorMessage)
        End If

        sw.Close()
    End Sub

    Public Sub LogMessage(ByVal Logfilename As String, ByVal LogMessagePath As String, ByVal Message As String, Optional ByVal infoPrefix As String = "")
        Dim Log_Filename As String = LogMessagePath & Logfilename
        Dim sw As StreamWriter = AppendText(Log_Filename)

        If Len(infoPrefix) > 0 Then
            sw.WriteLine(Now() & " " & "MESSAGE: " & Message & "," & infoPrefix & "," & Message)
        Else
            sw.WriteLine(Now() & " " & "MESSAGE: " & Message)
        End If
        'MsgBox("Log Pathname:" & LogMessagePath)
        sw.Close()
    End Sub
End Class
