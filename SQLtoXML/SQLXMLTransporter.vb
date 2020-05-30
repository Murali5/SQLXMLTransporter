Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml

Public Class SQLXMLTransporter

    Private _dataTableName As String()
    Public Property DataTableNames() As String()
        Get
            Return _dataTableName
        End Get
        Set(ByVal value As String())
            _dataTableName = value
        End Set
    End Property

    Private _dataBaseTableColumn As Object
    Public Property DataBaseTableColumns() As Object
        Get
            Return _dataBaseTableColumn
        End Get
        Set(ByVal value As Object)
            _dataBaseTableColumn = value
        End Set
    End Property

    Property SQLData As sqlComponent

    ''' <summary>
    ''' Create a new instance of the SQL - XML Transporter
    ''' </summary>
    ''' <param name="conStr">Connection string to the DB</param>
    ''' <param name="dataTableNames">All datatable names in the database</param>
    ''' <param name="databaseColumnsName">All datattable coloums as array of array strings</param>
    ''' <remarks></remarks>
    Public Sub New(conStr As String, dataTableNames As String(), databaseColumnsName As Object)
        Me.SQLData = New sqlComponent(conStr)
        Me.DataBaseTableColumns = databaseColumnsName
        Me.DataTableNames = dataTableNames
    End Sub

    Private Sub ExportXML(ByVal XMLFileName As String)
        Dim myDoc As XmlDocument = New XmlDocument
        Dim myXmlSettings As XmlWriterSettings = New XmlWriterSettings

        myXmlSettings.Indent = True
        myXmlSettings.Encoding = System.Text.Encoding.UTF32
        Dim myXml As XmlWriter = XmlWriter.Create(XMLFileName, myXmlSettings)

        myXml.WriteStartDocument()
        myXml.WriteStartElement("DataBaseTables")

        For i = 0 To DataTableNames.Length - 1
            Dim CurDataTableName As String = DataTableNames(i)
            Dim CurTableData = SQLData.GetTableQurery("SELECT * FROM " & CurDataTableName)
            myXml.WriteStartElement(CurDataTableName)

            If Not CurTableData.Rows.Count = 0 Then
                For j = CurTableData.Rows.Count - 1 To 0 Step -1
                    myXml.WriteStartElement("ID" & (j + 1).ToString())
                    For k = 0 To CurTableData.Columns.Count - 1

                        Try
                            myXml.WriteElementString(DataBaseTableColumns(i)(k).ToString, IIf(IsDBNull(CurTableData.Rows(j).Item(k)), "DBNull", CurTableData.Rows(j).Item(k).ToString))
                        Catch ex As Exception
                            '   logwriteline(ex.Message & " " & ex.StackTrace)
                            LogWriteLine("XML 1:" & ex.Message & " - " & ex.StackTrace)
                        End Try

                    Next
                    myXml.WriteEndElement()
                Next
            End If
            myXml.WriteEndElement()
        Next
        myXml.WriteEndElement()
        myXml.WriteEndDocument()
        myXml.Close()
    End Sub

    Private Sub ImportXML(ByVal FileName As String)
        Dim myXMLReader As XmlReader = XmlReader.Create(FileName)
        Dim myXMLexp As XMLSQL = New XMLSQL(DataBaseTableColumns, DataTableNames)

        While myXMLReader.Read()

            Select Case myXMLReader.NodeType
                Case XmlNodeType.Element
                    myXMLexp.PickInside(myXMLReader.Name, Me.SQLData)
                Case XmlNodeType.Text
                    myXMLexp.AddColumnValue(IIf(myXMLReader.Value = "DBNull", "", myXMLReader.Value))
                Case XmlNodeType.EndElement
                    myXMLexp.PickOutside(Me.SQLData)
            End Select
        End While
        myXMLReader.Close()
        myXMLexp = Nothing
        myXMLReader = Nothing

    End Sub



End Class
