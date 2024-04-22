'Imports System.Text
'Imports System.IO

'Imports NPOI.HSSF.UserModel
'Imports System.ComponentModel

'Namespace Apchy.Helper

'    Public Class ExcelHelper
'        Public Shared ListColumnsName As System.Collections.SortedList
'        Public Shared RowLimited As Integer = 65000

'        'Row limits older excel verion per sheet, the row limit for excel 2003 is 65536
'        'Const rowLimit As Integer = 65000


'        Public Shared Sub SaveFile(ByVal DataTable As DataTable, ByVal FileName As String)
'            Dim myFileStream As FileStream = Nothing

'            Try
'                myFileStream = New FileStream(FileName, FileMode.Create)

'                Dim myWorkBook As HSSFWorkbook = CreateWorkbook(DataTable)

'                myWorkBook.Write(myFileStream)

'            Catch ex As Exception
'            Finally
'                If myFileStream IsNot Nothing Then
'                    myFileStream.Close()
'                End If
'            End Try
'        End Sub

'        Protected Shared Function CreateWorkbook(ByVal ListSource As IListSource, ByVal ColumnHeader As System.Collections.SortedList) As HSSFWorkbook
'            Dim myWorkbook As New HSSFWorkbook

'            Dim RowCount As Integer = 0
'            Dim SheetCount As Integer = 1

'            Dim newSheet As HSSFSheet = Nothing

'            newSheet = myWorkbook.CreateSheet("Sheet1" & SheetCount)

'            CreateColumnHeader(newSheet, ColumnHeader)

'            For Each DataRow In ListSource.GetList()
'                RowCount += 1

'                If RowCount = RowLimited Then
'                    RowCount = 1
'                    SheetCount += 1
'                    newSheet = myWorkbook.CreateSheet("Sheet" & SheetCount)
'                    CreateColumnHeader(newSheet, ColumnHeader)
'                End If

'                Dim newExcelRow As HSSFRow = newSheet.CreateRow(RowCount)

'                Dim I As Integer
'                Dim myFieldType As Type
'                Dim myFieldValue As Object

'                For I = 0 To ListColumnsName.Count - 1
'                    Dim myFieldName As String = ListColumnsName.GetKey(I)

'                    myFieldType = DataRow.Item(myFieldName).GetType
'                    myFieldValue = DataRow.Item(myFieldName)

'                    Dim newcell As HSSFCell = newExcelRow.CreateCell(I)
'                    InsertCell(newcell, myFieldType, myFieldValue)
'                Next
'            Next
'        End Function

'        Protected Shared Function CreateWorkbook(ByVal Datatable As DataTable) As HSSFWorkbook
'            Dim myWorkbook As HSSFWorkbook

'            myWorkbook = New HSSFWorkbook

'            Dim RowCount As Integer = 0
'            Dim SheetCount As Integer = 1

'            Dim newSheet As HSSFSheet = Nothing

'            newSheet = myWorkbook.CreateSheet("Sheet1" & SheetCount)

'            CreateColumnHeader(newSheet, Datatable)
'            For Each DataRow As DataRow In Datatable.Rows
'                RowCount += 1

'                If RowCount = RowLimited Then
'                    RowCount = 1
'                    SheetCount += 1
'                    newSheet = myWorkbook.CreateSheet("Sheet" & SheetCount)
'                    CreateColumnHeader(newSheet, Datatable)
'                End If

'                Dim newExcelRow As HSSFRow = newSheet.CreateRow(RowCount)

'                Dim I As Integer
'                Dim myFieldType As Type
'                Dim myFieldValue As Object

'                If ListColumnsName Is Nothing OrElse ListColumnsName.Count = 0 Then
'                    For I = 0 To Datatable.Columns.Count - 1
'                        myFieldType = DataRow.Item(I).GetType
'                        myFieldValue = DataRow.Item(I)

'                        Dim newcell As HSSFCell = newExcelRow.CreateCell(I)
'                        InsertCell(newcell, myFieldType, myFieldValue)
'                    Next
'                Else
'                    For I = 0 To ListColumnsName.Count - 1
'                        Dim myFieldName As String = ListColumnsName.GetKey(I)
'                        myFieldType = DataRow.Item(myFieldName).GetType
'                        myFieldValue = DataRow.Item(myFieldName)

'                        Dim newcell As HSSFCell = newExcelRow.CreateCell(I)
'                        InsertCell(newcell, myFieldType, myFieldValue)
'                    Next
'                End If
'            Next
'        End Function

'        Protected Shared Sub InsertCell(ByRef newCell As HSSFCell, ByVal FieldType As Type, ByVal FieldValue As Object)

'            Select Case FieldType.ToString
'                Case "System.String"
'                    '字符串类型
'                    FieldValue = FieldValue.Replace("&", "&")
'                    FieldValue = FieldValue.Replace(">", ">")
'                    FieldValue = FieldValue.Replace("<", "<")
'                    newCell.SetCellValue(FieldValue)
'                    Exit Select
'                Case "System.DateTime"
'                    '日期类型
'                    Dim dateV As DateTime
'                    DateTime.TryParse(FieldValue, dateV)
'                    newCell.SetCellValue(dateV)

'                    '格式化显示
'                    Dim myWorkbook As HSSFWorkbook = newCell.Sheet.Workbook

'                    Dim cellStyle As HSSFCellStyle = myWorkbook.CreateCellStyle()
'                    Dim format As HSSFDataFormat = myWorkbook.CreateDataFormat()
'                    cellStyle.DataFormat = format.GetFormat("yyyy-mm-dd")
'                    newCell.CellStyle = cellStyle

'                    Exit Select
'                Case "System.Boolean"
'                    '布尔型
'                    Dim boolV As Boolean = False
'                    Boolean.TryParse(FieldValue, boolV)
'                    newCell.SetCellValue(boolV)
'                    Exit Select
'                    '整型
'                Case "System.Int16", "System.Int32", "System.Int64", "System.Byte"
'                    Dim intV As Integer = 0
'                    Integer.TryParse(FieldValue, intV)
'                    newCell.SetCellValue(intV.ToString())
'                    Exit Select
'                    '浮点型
'                Case "System.Decimal", "System.Double"
'                    Dim doubV As Double = 0
'                    Double.TryParse(FieldValue, doubV)
'                    newCell.SetCellValue(doubV)
'                    Exit Select
'                Case "System.DBNull"
'                    '空值处理
'                    newCell.SetCellValue("")
'                    Exit Select
'                Case Else
'                    Throw (New Exception(FieldType.ToString() & "：类型数据无法处理!"))
'            End Select
'        End Sub

'        Protected Shared Sub CreateColumnHeader(ByRef WorkSheet As HSSFSheet, ByVal ColumnHeader As System.Collections.SortedList)
'            Dim CellIndex As Integer = 0

'            If ColumnHeader Is Nothing OrElse ColumnHeader.Count = 0 Then Exit Sub

'            For CellIndex = 0 To ColumnHeader.Count - 1
'                Dim myColumnName As String = ColumnHeader.GetKey(CellIndex)

'                Dim newRow As HSSFRow = WorkSheet.CreateRow(0)
'                Dim newCell As HSSFCell = newRow.CreateCell(CellIndex)

'                newCell.SetCellValue(myColumnName)
'            Next
'        End Sub

'        Protected Shared Sub CreateColumnHeader(ByRef WorkSheet As HSSFSheet, ByVal DataTable As DataTable)
'            Dim CellIndex As Integer = 0

'            If ListColumnsName Is Nothing OrElse ListColumnsName.Count = 0 Then
'                For I = 0 To DataTable.Columns.Count - 1
'                    Dim myColumnName As String = DataTable.Columns(I).ColumnName

'                    Dim newRow As HSSFRow = WorkSheet.CreateRow(0)
'                    Dim newCell As HSSFCell = newRow.CreateCell(CellIndex)

'                    newCell.SetCellValue(myColumnName)
'                    CellIndex += 1
'                Next
'            Else
'                For Each obj As DictionaryEntry In ListColumnsName
'                    Dim newRow As HSSFRow = WorkSheet.CreateRow(0)
'                    Dim newCell As HSSFCell = newRow.CreateCell(CellIndex)

'                    newCell.SetCellValue(obj.Value.ToString)

'                    CellIndex += 1
'                Next
'            End If
'        End Sub

'        Public Class NoSort
'            Implements System.Collections.IComparer

'            Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
'                Return -1
'            End Function
'        End Class

'        'Row limits older excel verion per sheet, the row limit for excel 2003 is 65536
'        Const rowLimit As Integer = 65000

'        Public Shared Sub WriteExcelXml(ByVal ToStream As Stream, ByVal dtInput As DataTable, ByVal filename As String)
'            Dim ds = New DataSet()

'            WriteExcelXml(ToStream, ds, filename)
'        End Sub

'        Public Shared Sub WriteExcelXml(ByVal ToStream As Stream, ByVal dsInput As DataSet, ByVal filename As String)
'            WriteWorkbookTemplateHeader(ToStream)
'            WriteWorksheets(ToStream, dsInput)
'            WriteWorkbookTemplateFooter(ToStream)
'        End Sub

'#Region "Excel XML Writting"
'        Private Shared Sub WriteWorkbookTemplateHeader(ByVal ToStream As IO.Stream)
'            Try
'                Dim sb = New StringBuilder(818)
'                sb.AppendFormat("<?xml version=""1.0""?>{0}", Environment.NewLine)
'                sb.AppendFormat("<?mso-application progid=""Excel.Sheet""?>{0}", Environment.NewLine)
'                sb.AppendFormat("<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet""{0}", Environment.NewLine)
'                sb.AppendFormat(" xmlns:o=""urn:schemas-microsoft-com:office:office""{0}", Environment.NewLine)
'                sb.AppendFormat(" xmlns:x=""urn:schemas-microsoft-com:office:excel""{0}", Environment.NewLine)
'                sb.AppendFormat(" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet""{0}", Environment.NewLine)
'                sb.AppendFormat(" xmlns:html=""http://www.w3.org/TR/REC-html40"">{0}", Environment.NewLine)
'                sb.AppendFormat(" <Styles>{0}", Environment.NewLine)
'                sb.AppendFormat("  <Style ss:ID=""Default"" ss:Name=""Normal"">{0}", Environment.NewLine)
'                sb.AppendFormat("   <Alignment ss:Vertical=""Bottom""/>{0}", Environment.NewLine)
'                sb.AppendFormat("   <Borders/>{0}", Environment.NewLine)
'                sb.AppendFormat("   <Font ss:FontName=""Calibri"" x:Family=""Swiss"" ss:Size=""11"" ss:Color=""#000000""/>{0}", Environment.NewLine)
'                sb.AppendFormat("   <Interior/>{0}", Environment.NewLine)
'                sb.AppendFormat("   <NumberFormat/>{0}", Environment.NewLine)
'                sb.AppendFormat("   <Protection/>{0}", Environment.NewLine)
'                sb.AppendFormat("  </Style>{0}", Environment.NewLine)
'                sb.AppendFormat("  <Style ss:ID=""s62"">{0}", Environment.NewLine)
'                sb.AppendFormat("   <Font ss:FontName=""Calibri"" x:Family=""Swiss"" ss:Size=""11"" ss:Color=""#000000""{0}", Environment.NewLine)
'                sb.AppendFormat("    ss:Bold=""1""/>{0}", Environment.NewLine)
'                sb.AppendFormat("  </Style>{0}", Environment.NewLine)
'                sb.AppendFormat("  <Style ss:ID=""s63"">{0}", Environment.NewLine)
'                sb.AppendFormat("   <NumberFormat ss:Format=""Short Date""/>{0}", Environment.NewLine)
'                sb.AppendFormat("  </Style>{0}", Environment.NewLine)
'                sb.AppendFormat(" </Styles>{0}", Environment.NewLine)
'                'sb.Append("{0}\r\n</Workbook>")

'                Dim Output() As Byte = GetBytes(sb.ToString())

'                ToStream.Write(Output, 0, Output.Length)
'            Catch ex As Exception
'                Throw ex
'            End Try
'        End Sub

'        Private Shared Sub WriteWorkbookTemplateFooter(ByVal ToStream As IO.Stream)
'            Try
'                Dim myString As String = "\r\n</Workbook>"
'                Dim Output() As Byte = GetBytes(myString)
'                ToStream.Write(Output, 0, Output.Length)
'            Catch ex As Exception
'                Throw ex
'            End Try
'        End Sub

'        Private Shared Sub WriteWorksheets(ByVal ToStream As Stream, ByVal source As DataSet)
'            Try
'                Dim myString As String = ""
'                If source Is Nothing OrElse source.Tables.Count = 0 Then
'                    myString = "<Worksheet ss:Name=""Sheet1"">" & vbCr & vbLf & "<Table>" & vbCr & vbLf & "<Row><Cell><Data ss:Type=""String""></Data></Cell></Row>" & vbCr & vbLf & "</Table>" & vbCr & vbLf & "</Worksheet>"
'                    Dim Output() As Byte = GetBytes(myString)
'                    ToStream.Write(Output, 0, Output.Length)
'                End If

'                For Each dt As DataTable In source.Tables
'                    If dt.Rows.Count = 0 Then
'                        myString = "<Worksheet ss:Name=""" & replaceXmlChar(dt.TableName) & """>" & vbCr & vbLf & "<Table>" & vbCr & vbLf & "<Row><Cell  ss:StyleID=""s62""><Data ss:Type=""String""></Data></Cell></Row>" & vbCr & vbLf & "</Table>" & vbCr & vbLf & "</Worksheet>"

'                        Dim Output() As Byte = GetBytes(myString)
'                        ToStream.Write(Output, 0, Output.Length)
'                    Else
'                        'write each row data                
'                        Dim sheetCount = 0
'                        For i As Integer = 0 To dt.Rows.Count - 1
'                            Dim sw As New StringWriter
'                            If (i Mod rowLimit) = 0 Then
'                                'add close tags for previous sheet of the same data table
'                                If (i \ rowLimit) > sheetCount Then
'                                    sw.Write(vbCr & vbLf & "</Table>" & vbCr & vbLf & "</Worksheet>")
'                                    sheetCount = (i \ rowLimit)
'                                End If
'                                sw.Write(vbCr & vbLf & "<Worksheet ss:Name=""" & replaceXmlChar(dt.TableName) & (If(((i \ rowLimit) = 0), "", Convert.ToString(i \ rowLimit))) & """>" & vbCr & vbLf & "<Table>")
'                                'write column name row
'                                sw.Write(vbCr & vbLf & "<Row>")
'                                For Each dc As DataColumn In dt.Columns
'                                    sw.Write(String.Format("<Cell ss:StyleID=""s62""><Data ss:Type=""String"">{0}</Data></Cell>", replaceXmlChar(dc.ColumnName)))
'                                Next
'                                sw.Write("</Row>")
'                            End If
'                            sw.Write(vbCr & vbLf & "<Row>")
'                            For Each dc As DataColumn In dt.Columns
'                                sw.Write(getCell(dc.DataType, dt.Rows(i)(dc.ColumnName)))
'                            Next
'                            sw.Write("</Row>")

'                            Dim myOutput() As Byte = GetBytes(sw.ToString)
'                            ToStream.Write(myOutput, 0, myOutput.Length)
'                        Next

'                        myString = vbCr & vbLf & "</Table>" & vbCr & vbLf & "</Worksheet>"
'                        Dim Output() As Byte = GetBytes(myString)
'                        ToStream.Write(Output, 0, Output.Length)
'                    End If
'                Next
'            Catch ex As Exception
'                Throw ex
'            End Try
'        End Sub

'        Private Shared Sub WriteCell(ByVal ToStream As Stream, ByVal type As Type, ByVal cellData As Object)
'            Try

'                Dim data = If((TypeOf cellData Is DBNull), "", cellData)

'                If type.Name.Contains("Int") OrElse type.Name.Contains("Double") OrElse type.Name.Contains("Decimal") Then
'                    Dim Output() As Byte = GetBytes(String.Format("<Cell><Data ss:Type=""Number"">{0}</Data></Cell>", data))

'                    ToStream.Write(Output, 0, Output.Length)
'                End If

'                If type.Name.Contains("Date") AndAlso data.ToString() <> String.Empty Then
'                    Dim Output() As Byte = GetBytes(String.Format("<Cell ss:StyleID=""s63""><Data ss:Type=""DateTime"">{0}</Data></Cell>", Convert.ToDateTime(data).ToString("yyyy-MM-dd")))

'                    ToStream.Write(Output, 0, Output.Length)
'                End If
'            Catch ex As Exception
'                Throw ex
'            End Try
'        End Sub
'#End Region


'#Region "Utility"
'        Private Shared Function GetBytes(ByVal InputString As String) As Byte()
'            Return Encoding.UTF8.GetBytes(InputString)
'        End Function

'        Private Shared Function replaceXmlChar(ByVal input As String) As String
'            input = input.Replace("&", "&amp")
'            input = input.Replace("<", "&lt;")
'            input = input.Replace(">", "&gt;")
'            input = input.Replace("""", "&quot;")
'            input = input.Replace("'", "&apos;")
'            Return input
'        End Function
'#End Region

'#Region "Old Version"
'        Private Shared Function getWorkbookTemplate() As String
'            Dim sb = New StringBuilder(818)
'            sb.AppendFormat("<?xml version=""1.0""?>{0}", Environment.NewLine)
'            sb.AppendFormat("<?mso-application progid=""Excel.Sheet""?>{0}", Environment.NewLine)
'            sb.AppendFormat("<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet""{0}", Environment.NewLine)
'            sb.AppendFormat(" xmlns:o=""urn:schemas-microsoft-com:office:office""{0}", Environment.NewLine)
'            sb.AppendFormat(" xmlns:x=""urn:schemas-microsoft-com:office:excel""{0}", Environment.NewLine)
'            sb.AppendFormat(" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet""{0}", Environment.NewLine)
'            sb.AppendFormat(" xmlns:html=""http://www.w3.org/TR/REC-html40"">{0}", Environment.NewLine)
'            sb.AppendFormat(" <Styles>{0}", Environment.NewLine)
'            sb.AppendFormat("  <Style ss:ID=""Default"" ss:Name=""Normal"">{0}", Environment.NewLine)
'            sb.AppendFormat("   <Alignment ss:Vertical=""Bottom""/>{0}", Environment.NewLine)
'            sb.AppendFormat("   <Borders/>{0}", Environment.NewLine)
'            sb.AppendFormat("   <Font ss:FontName=""Calibri"" x:Family=""Swiss"" ss:Size=""11"" ss:Color=""#000000""/>{0}", Environment.NewLine)
'            sb.AppendFormat("   <Interior/>{0}", Environment.NewLine)
'            sb.AppendFormat("   <NumberFormat/>{0}", Environment.NewLine)
'            sb.AppendFormat("   <Protection/>{0}", Environment.NewLine)
'            sb.AppendFormat("  </Style>{0}", Environment.NewLine)
'            sb.AppendFormat("  <Style ss:ID=""s62"">{0}", Environment.NewLine)
'            sb.AppendFormat("   <Font ss:FontName=""Calibri"" x:Family=""Swiss"" ss:Size=""11"" ss:Color=""#000000""{0}", Environment.NewLine)
'            sb.AppendFormat("    ss:Bold=""1""/>{0}", Environment.NewLine)
'            sb.AppendFormat("  </Style>{0}", Environment.NewLine)
'            sb.AppendFormat("  <Style ss:ID=""s63"">{0}", Environment.NewLine)
'            sb.AppendFormat("   <NumberFormat ss:Format=""Short Date""/>{0}", Environment.NewLine)
'            sb.AppendFormat("  </Style>{0}", Environment.NewLine)
'            sb.AppendFormat(" </Styles>{0}", Environment.NewLine)
'            sb.Append("{0}\r\n</Workbook>")
'            Return sb.ToString()
'        End Function

'        Private Shared Function getCell(ByVal type As Type, ByVal cellData As Object) As String
'            Dim data = If((TypeOf cellData Is DBNull), "", cellData)
'            If type.Name.Contains("Int") OrElse type.Name.Contains("Double") OrElse type.Name.Contains("Decimal") Then
'                Return String.Format("<Cell><Data ss:Type=""Number"">{0}</Data></Cell>", data)
'            End If
'            If type.Name.Contains("Date") AndAlso data.ToString() <> String.Empty Then
'                Return String.Format("<Cell ss:StyleID=""s63""><Data ss:Type=""DateTime"">{0}</Data></Cell>", Convert.ToDateTime(data).ToString("yyyy-MM-dd"))
'            End If
'            Return String.Format("<Cell><Data ss:Type=""String"">{0}</Data></Cell>", replaceXmlChar(data.ToString()))
'        End Function
'        Private Shared Function getWorksheets(ByVal source As DataSet) As String
'            Dim sw = New StringWriter()
'            If source Is Nothing OrElse source.Tables.Count = 0 Then
'                sw.Write("<Worksheet ss:Name=""Sheet1"">" & vbCr & vbLf & "<Table>" & vbCr & vbLf & "<Row><Cell><Data ss:Type=""String""></Data></Cell></Row>" & vbCr & vbLf & "</Table>" & vbCr & vbLf & "</Worksheet>")
'                Return sw.ToString()
'            End If
'            For Each dt As DataTable In source.Tables
'                If dt.Rows.Count = 0 Then
'                    sw.Write("<Worksheet ss:Name=""" & replaceXmlChar(dt.TableName) & """>" & vbCr & vbLf & "<Table>" & vbCr & vbLf & "<Row><Cell  ss:StyleID=""s62""><Data ss:Type=""String""></Data></Cell></Row>" & vbCr & vbLf & "</Table>" & vbCr & vbLf & "</Worksheet>")
'                Else
'                    'write each row data                
'                    Dim sheetCount = 0
'                    For i As Integer = 0 To dt.Rows.Count - 1
'                        If (i Mod rowLimit) = 0 Then
'                            'add close tags for previous sheet of the same data table
'                            If (i \ rowLimit) > sheetCount Then
'                                sw.Write(vbCr & vbLf & "</Table>" & vbCr & vbLf & "</Worksheet>")
'                                sheetCount = (i \ rowLimit)
'                            End If
'                            sw.Write(vbCr & vbLf & "<Worksheet ss:Name=""" & replaceXmlChar(dt.TableName) & (If(((i \ rowLimit) = 0), "", Convert.ToString(i \ rowLimit))) & """>" & vbCr & vbLf & "<Table>")
'                            'write column name row
'                            sw.Write(vbCr & vbLf & "<Row>")
'                            For Each dc As DataColumn In dt.Columns
'                                sw.Write(String.Format("<Cell ss:StyleID=""s62""><Data ss:Type=""String"">{0}</Data></Cell>", replaceXmlChar(dc.ColumnName)))
'                            Next
'                            sw.Write("</Row>")
'                        End If
'                        sw.Write(vbCr & vbLf & "<Row>")
'                        For Each dc As DataColumn In dt.Columns
'                            sw.Write(getCell(dc.DataType, dt.Rows(i)(dc.ColumnName)))
'                        Next
'                        sw.Write("</Row>")
'                    Next
'                    sw.Write(vbCr & vbLf & "</Table>" & vbCr & vbLf & "</Worksheet>")
'                End If
'            Next

'            Return sw.ToString()
'        End Function

'        Public Shared Function GetExcelXml(ByVal dtInput As DataTable, ByVal filename As String) As String
'            Dim excelTemplate = getWorkbookTemplate()
'            Dim ds = New DataSet()
'            ds.Tables.Add(dtInput.Copy())
'            Dim worksheets = getWorksheets(ds)
'            Dim excelXml = String.Format(excelTemplate, worksheets)
'            Return excelXml
'        End Function

'        Public Shared Function GetExcelXml(ByVal dsInput As DataSet, ByVal filename As String) As String
'            Dim excelTemplate = getWorkbookTemplate()
'            Dim worksheets = getWorksheets(dsInput)
'            Dim excelXml = String.Format(excelTemplate, worksheets)
'            Return excelXml
'        End Function

'#End Region


'        'Public Shared Sub ToExcel(ByVal dsInput As DataSet, ByVal filename As String, ByVal response As HttpResponse)
'        '    Dim excelXml = GetExcelXml(dsInput, filename)
'        '    response.Clear()
'        '    response.AppendHeader("Content-Type", "application/vnd.ms-excel")
'        '    response.AppendHeader("Content-disposition", "attachment; filename=" & filename)
'        '    response.Write(excelXml)
'        '    response.Flush()
'        '    response.[End]()
'        'End Sub

'        'Public Shared Sub ToExcel(ByVal dtInput As DataTable, ByVal filename As String, ByVal response As HttpResponse)
'        '    Dim ds = New DataSet()
'        '    ds.Tables.Add(dtInput.Copy())
'        '    ToExcel(ds, filename, response)
'        'End Sub

'    End Class

'End Namespace