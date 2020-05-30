# SQLXMLTransporter
Imports or Exports the data from SQL to XML.
This is my first open source project to the community.

Language used: VB.NET
Softwares used:
Visual studio 2012 Ultimate
SQL compact version 4.0

How to use:
Sample code is given below
 Dim datatablenames As String() = New String() {"table1", "table2", "table3"}
        Dim databaseColumns As Object = {New String() {"tbl1Column1", "tbl1Column2", "tbl1Column3"},
                                         New String() {"tbl2Column1", "tbl2Column2", "tbl2Column3"},
                                          New String() {"tbl3Column1", "tbl3Column2", "tbl3Column3"}}
        Dim sqlXMLTrans As New SQLtoXML.SQLXMLTransporter("Server=localhost\sqlexpress", datatablenames, databaseColumns)
