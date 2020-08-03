Imports System.IO
Imports System.Xml

Public Class xmlClassConverter

    Public Function CreateXltoXML(dt As DataTable, path As String, RowName As String) As Boolean

        Dim i As Integer = 0
        Dim IsCreated As Boolean = False
        Dim exMessage As String = Nothing
        Try
            Dim writer As XmlTextWriter = New XmlTextWriter(path, System.Text.Encoding.UTF8)
            writer.WriteStartDocument(True)
            writer.Formatting = Formatting.Indented
            writer.Indentation = 2
            writer.WriteStartElement(RowName)


            Dim ColumnNames As List(Of String) = dt.Columns.Cast(Of DataColumn)().ToList().Select(Function(x) x.ColumnName).ToList() 'Column Names List  
            Dim RowList As List(Of DataRow) = dt.Rows.Cast(Of DataRow)().ToList()

            For Each dw As DataRow In RowList
                For Each item As String In ColumnNames
                    writer.WriteStartElement(item)
                    writer.WriteString(dw.ItemArray(i).ToString())
                    writer.WriteEndElement()
                    i += 1
                Next
                i = 0
            Next

            writer.WriteEndElement()
            writer.WriteEndDocument()
            writer.Close()

            If (File.Exists(path)) Then
                IsCreated = True
            End If

        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try

        Return IsCreated
    End Function




End Class
