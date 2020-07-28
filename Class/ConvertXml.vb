Imports System.IO
Imports System.Text
Imports System.Xml
Imports DocumentFormat.OpenXml.Office2010.ExcelAc

Public Class ConvertXml


    Public Function CreateXltoXML(dt As DataTable, XmlFile As String, RowName As String) As Boolean
        Dim exMessage As String = " "
        Dim IsCreated As Boolean = False
        Try

            'Dim dt As DataTable = GetTableDataXl(XlFile)
            'Dim writer As XmlTextWriter = New XmlTextWriter(XmlFile, System.Text.Encoding.UTF8)
            Dim settings As XmlWriterSettings = New XmlWriterSettings()
            settings.Indent = True
            settings.Encoding = Encoding.Unicode


            'Using writer As XmlTextWriter = XmlTextWriter.Create(XmlFile, settings)
            Using writer As XmlTextWriter = New XmlTextWriter(XmlFile, System.Text.Encoding.UTF8)
                writer.Formatting = Formatting.Indented
                writer.WriteStartDocument(True)
                writer.Formatting = Formatting.Indented
                writer.Indentation = 2
                writer.WriteStartElement(RowName)

                Dim i As Integer = 0
                Dim ColumnNames As List(Of String) = dt.Columns.Cast(Of DataColumn)().ToList().Select(Function(x) x.ColumnName).ToList() 'Column Names List  
                Dim RowList As List(Of DataRow) = dt.Rows.Cast(Of DataRow)().ToList()
                For Each dr As DataRow In RowList
                    'writer.WriteStartElement(RowName)
                    For Each str As String In ColumnNames
                        writer.WriteStartElement(str)
                        writer.WriteString(dr.ItemArray(i).ToString())
                        writer.WriteEndElement()
                        i += 1
                    Next
                    'writer.WriteEndElement()
                    i = 0
                Next

                writer.WriteEndElement()
                writer.WriteEndDocument()
                writer.Flush()
                writer.Close()
                writer.Dispose()
                If (File.Exists(XmlFile)) Then
                    IsCreated = True
                End If
            End Using

            GC.Collect()
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try

        Return IsCreated

    End Function

End Class
