Imports System
Imports System.Data
Imports System.Data.SqlServerCe

Public Class sqlComponent
    Dim sqlCon As SqlCeConnection
    Dim sqlCmd As SqlCeCommand
    Dim sqlDS As DataSet
    Dim sqlDA As SqlCeDataAdapter

    Private _Constr As String

    Public ReadOnly Property ConStr() As String
        Get
            Return _Constr
        End Get
    End Property

    Public Sub New(Constring As String)
        Me._Constr = Constring
        sqlCon = New SqlCeConnection()
        sqlCmd = New SqlCeCommand
        sqlDS = New DataSet
        sqlDA = New SqlCeDataAdapter

    End Sub

    Public Sub ExecuteQurery(ByVal cmdTxt As String)
        Try
            sqlCon = New SqlCeConnection(ConStr)
            sqlCmd = New SqlCeCommand(cmdTxt, sqlCon)
            If Not sqlCon.State = ConnectionState.Open Then sqlCon.Open()
            sqlCmd.ExecuteNonQuery()
        Catch ex1 As SqlCeException
            LogWriteLine(ex1.Message & " " & ex1.StackTrace)
        Catch ex As Exception
            logwriteline(ex.Message & " " & ex.StackTrace)

        Finally
            If Not sqlCon.State = ConnectionState.Closed Then sqlCon.Close()

        End Try
    End Sub


    Public Function GetTableQurery(ByVal cmdTxt As String) As DataTable
        Try
            sqlCon = New SqlCeConnection(ConStr)
            sqlCmd = New SqlCeCommand(cmdTxt, sqlCon)

            sqlDA = New SqlCeDataAdapter(sqlCmd)
            sqlDS = New DataSet
            sqlDA.Fill(sqlDS)

            Return sqlDS.Tables(0)

        Catch ex As Exception

            logwriteline(ex.Message & " " & ex.StackTrace)
            Return New DataTable
        Finally
            If Not sqlCon.State = ConnectionState.Closed Then sqlCon.Close()

        End Try
    End Function

End Class
