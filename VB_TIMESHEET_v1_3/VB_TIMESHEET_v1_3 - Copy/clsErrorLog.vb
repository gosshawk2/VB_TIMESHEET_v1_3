Imports System.IO
Imports System.IO.File
Public Class clsErrorLog
    Public Sub LogError(ByVal Logfilename As String, ByVal LogErrorPath As String, ByVal ErrorMessage As String, Optional ByVal RoutineName As String = "",
                        Optional ByVal infoPrefix As String = "",
                        Optional ByVal usernamePrefix As String = "")
        Dim Log_Filename As String


        Dim sw As StreamWriter

        If len(usernamePrefix) > 0 Then
            Log_Filename = LogErrorPath & "\Errors\" & Logfilename & "_" & usernamePrefix
        Else
            Log_Filename = LogErrorPath & "\Errors\" & Logfilename
        End If
        sw = AppendText(Log_Filename)
        If Len(infoPrefix) > 0 Then
            sw.WriteLine(Now() & " " & infoPrefix & ", {" & RoutineName & "} ," & ErrorMessage)
        Else
            sw.WriteLine(Now() & " {" & RoutineName & "} ," & ErrorMessage)
        End If

        sw.Close()
    End Sub

    Public Sub LogMessage(ByVal Logfilename As String, ByVal LogMessagePath As String, ByVal Message As String,
                          Optional ByVal infoPrefix As String = "",
                          Optional ByVal usernamePrefix As String = "")
        Dim Log_Filename As String


        If len(usernamePrefix) > 0 Then
            Log_Filename = LogMessagePath & "\Messages\" & Logfilename & "_" & usernamePrefix
            'Log_Filename = LogMessagePath & "\Messages\" & Logfilename
        Else
            Log_Filename = LogMessagePath & "\Messages\" & Logfilename
        End If

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
