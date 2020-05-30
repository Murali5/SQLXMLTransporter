Imports System
Imports System.IO

Module Module1
    Public LogMsg As String = ""
    Private _writer As StreamWriter
    Private LineNo As Integer = 0

    Sub LogWriteLine(ByVal logText As String, Optional ByVal Append As Boolean = True)
        Try
            If Not File.Exists("C:\FLSmidth_Log\logFile.txt") Then
                If Not Directory.Exists("C:\FLSmidth_Log\") Then
                    Directory.CreateDirectory("C:\FLSmidth_Log\")
                End If
                File.Create("C:\FLSmidth_Log\logFile.txt")
            End If

            LineNo += 1
            _writer = My.Computer.FileSystem.OpenTextFileWriter("C:\FLSmidth_Log\logFile.txt", Append)
            _writer.WriteLine(LineNo.ToString + " " + Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + "   " + logText)
            _writer.Close()
        Catch ex As Exception
            LogWriteLine(ex.Message & " " & ex.StackTrace)
        End Try

    End Sub


End Module
