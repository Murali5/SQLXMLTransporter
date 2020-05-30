Imports System
Imports System.Xml
Imports System.IO

Public Class XMLSQL
    Private _ValueGained As Boolean
    Private _CurPickedTable As Integer
    Public Property CurrentlyPickedTable() As Integer
        Get
            Return _CurPickedTable
        End Get
        Set(ByVal value As Integer)
            _CurPickedTable = value
        End Set
    End Property

    Private _curColumnToPick As Integer
    Public Property CurrentColumnToPick() As Integer
        Get
            Return _curColumnToPick
        End Get
        Set(ByVal value As Integer)
            _curColumnToPick = value
        End Set
    End Property

    Private _PickFeature As ElementToPick_e
    Public Property PickFeature() As ElementToPick_e
        Get
            Return _PickFeature
        End Get
        Set(ByVal value As ElementToPick_e)
            _PickFeature = value
        End Set
    End Property

    Private _CurRow As String()
    Public Property CurrentRow() As String()
        Get
            Return _CurRow
        End Get
        Set(ByVal value As String())
            _CurRow = value
        End Set
    End Property

    Private _databaseTableColumn As Object()
    Public Property DatabaseTableColumn() As Object()
        Get
            Return _databaseTableColumn
        End Get
        Set(ByVal value As Object())
            _databaseTableColumn = value
        End Set
    End Property

    Private _datatablename As String()
    Public Property DataTableName() As String()
        Get
            Return _datatablename
        End Get
        Set(ByVal value As String())
            _datatablename = value
        End Set
    End Property

    Sub New(databasecoumn As Object(), tablenames As String())
        DatabaseTableColumn = databasecoumn
        PickFeature = ElementToPick_e.PickNothing
        DataTableName = tablenames
    End Sub

    Sub AddColumnValue(ByVal Value As String)
        _ValueGained = True
        _CurRow(_curColumnToPick) = Value
    End Sub

    Sub PickInside(ByVal value As String, ByVal sql As sqlComponent)
        If PickFeature = ElementToPick_e.Row Then
            _ValueGained = False
        End If

        If PickFeature = ElementToPick_e.ColumnName Then
            If Not _ValueGained Then
                PickOutside(sql)
            End If
        End If
        PickFeature = CType(CInt(PickFeature) + 1, ElementToPick_e)

        If PickFeature = ElementToPick_e.Table Then
            NewTable(value)
        ElseIf PickFeature = ElementToPick_e.Row Then
            ClearRowData()
        ElseIf PickFeature = ElementToPick_e.ColumnName Then
            For i = 0 To DatabaseTableColumn(CurrentlyPickedTable).length - 1
                If DatabaseTableColumn(CurrentlyPickedTable)(i) = value Then
                    CurrentColumnToPick = i
                    Exit Sub
                End If
            Next
        End If
    End Sub

    Sub PickOutside(ByVal sql As sqlComponent)
        PickFeature = CType(CInt(PickFeature) - 1, ElementToPick_e)
        If PickFeature = ElementToPick_e.Table Then
            EndRow(sql)
        End If
    End Sub

    Sub EndRow(ByVal sql As sqlComponent)
        Dim CmdStr As String = "INSERT INTO " & DataTableName(CurrentlyPickedTable) & " ("
        For i = 0 To DatabaseTableColumn(CurrentlyPickedTable).length - 1
            CmdStr &= DatabaseTableColumn(CurrentlyPickedTable)(i).ToString
            If i <> DatabaseTableColumn(CurrentlyPickedTable).length - 1 Then
                CmdStr &= " , "
            End If
        Next
        CmdStr &= ") VALUES ("
        For i = 0 To DatabaseTableColumn(CurrentlyPickedTable).length - 1
            CmdStr &= "'" & CurrentRow(i) & "'"
            If Not i = DatabaseTableColumn(CurrentlyPickedTable).length - 1 Then
                CmdStr &= " , "
            Else
                CmdStr &= ")"
            End If
        Next
        sql.ExecuteQurery(CmdStr)
    End Sub

    Sub ClearRowData()
        For i = 0 To CurrentRow.Length - 1
            CurrentRow(i) = ""
        Next
    End Sub

    Sub NewTable(ByVal TableName As String)
        For i = 0 To DataTableName.Length - 1
            If TableName = DataTableName(i) Then
                CurrentRow = Nothing
                CurrentRow = DatabaseTableColumn(i).clone
                ClearRowData()
                CurrentlyPickedTable = i
                Exit Sub
            End If
        Next
    End Sub
End Class


Public Enum ElementToPick_e
    PickNothing = -1
    Root = 0
    Table = 1
    Row = 2
    ColumnName = 3
    Value = 4
End Enum