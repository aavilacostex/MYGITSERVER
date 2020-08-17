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
            Dim i As Integer = 0
            Dim ColumnNames As List(Of String) = dt.Columns.Cast(Of DataColumn)().ToList().Select(Function(x) x.ColumnName).ToList() 'Column Names List  
            Dim RowList As List(Of DataRow) = dt.Rows.Cast(Of DataRow)().ToList()

            Dim settings As XmlWriterSettings = New XmlWriterSettings()
            settings.Indent = True
            settings.IndentChars = ("    ")
            settings.CloseOutput = True
            settings.OmitXmlDeclaration = True
            Using writer1 As XmlWriter = XmlWriter.Create(XmlFile, settings)
                writer1.WriteStartElement(RowName)
                'writer1.Formatting = Formatting.Indented
                'writer1.Indentation = 2

                For Each dr As DataRow In RowList
                    'writer.WriteStartElement(RowName)
                    For Each str As String In ColumnNames
                        writer1.WriteStartElement(str)
                        writer1.WriteString(dr.ItemArray(i).ToString())
                        writer1.WriteEndElement()
                        i += 1
                    Next
                    'writer.WriteEndElement()
                    i = 0
                Next
                writer1.WriteEndElement()
                writer1.WriteEndDocument()
                writer1.Flush()
            End Using

            If (File.Exists(XmlFile)) Then
                IsCreated = True
            End If

            'Using writer As XmlTextWriter = XmlTextWriter.Create(XmlFile, settings)
            '    'Using writer As XmlTextWriter = New XmlTextWriter(XmlFile, System.Text.Encoding.UTF8)
            '    writer.Formatting = Formatting.Indented
            '    writer.WriteStartDocument(True)
            '    writer.Formatting = Formatting.Indented
            '    writer.Indentation = 2
            '    writer.WriteStartElement(RowName)


            '    For Each dr As DataRow In RowList
            '        'writer.WriteStartElement(RowName)
            '        For Each str As String In ColumnNames
            '            writer.WriteStartElement(str)
            '            writer.WriteString(dr.ItemArray(i).ToString())
            '            writer.WriteEndElement()
            '            i += 1
            '        Next
            '        'writer.WriteEndElement()
            '        i = 0
            '    Next

            '    writer.WriteEndElement()
            '    writer.WriteEndDocument()
            '    'writer.Flush()
            '    'writer.Close()
            '    'writer.Dispose()
            '    If (File.Exists(XmlFile)) Then
            '        IsCreated = True
            '    End If
            'End Using

            GC.Collect()
        Catch ex As Exception
            exMessage = ex.ToString + ". " + ex.Message + ". " + ex.ToString
        End Try

        Return IsCreated

    End Function

    'Public Sub Dispose()
    '    'Me.Close
    'End Sub

End Class
