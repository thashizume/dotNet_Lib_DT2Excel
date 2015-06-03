﻿Public Class dt2excel

    Public Function ToExcel(dt As System.Data.DataTable, fileName As String, sheetName As String)
        Dim _xlsWorkbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(fileName)




        Return Nothing
    End Function

    Public Function ToDataTable(fileName As String, sheetName As String, columns As Long, _
                                Optional headerRowNumber As Long = 1, Optional skipRow As Long = 1, _
                                Optional loadColumns As Long() = Nothing) As System.Data.DataTable

        '   Validate Excel File
        If (New System.IO.FileInfo(fileName).Exists) = False Then Throw New Exception("can not found file :" & fileName)
        If existSheet(SpreadsheetGear.Factory.GetWorkbook(fileName), sheetName) = False Then Throw New Exception("can not found Worksheet : " & sheetName)

        '   Define Variable Items
        Dim _dt As System.Data.DataTable = Me.generateDataTable( _
                SpreadsheetGear.Factory.GetWorkbook(fileName).Worksheets(sheetName), columns, headerRowNumber)

        '   Load Data
        Dim _sheet As SpreadsheetGear.IWorksheet = SpreadsheetGear.Factory.GetWorkbook(fileName).Worksheets(sheetName)
        Dim _r As Long = 1
        Dim _eorBuffer As String
        Do
            Dim row As System.Data.DataRow = _dt.NewRow
            _eorBuffer = String.Empty
            For i As Integer = 0 To columns - 1
                row(i) = _sheet.Cells(headerRowNumber - 1 + _r, i).Value
                _eorBuffer = _eorBuffer & _sheet.Cells(headerRowNumber - 1 + _r, i).Value
            Next
            If _eorBuffer = String.Empty Then Exit Do
            _dt.Rows.Add(row)
            _r = _r + 1
        Loop
        Return _dt
    End Function

    Private Function existSheet(xls As SpreadsheetGear.IWorkbook, sheetName As String) As Boolean
        For Each _sheet As SpreadsheetGear.IWorksheet In xls.Worksheets
            If _sheet.Name = sheetName Then Return True
        Next
        Return False
    End Function

    Private Function generateDataTable(sheet As SpreadsheetGear.IWorksheet, columns As Long, headerRowNumber As Long) As System.Data.DataTable

        Dim columnName As String() = getColumnName(sheet, columns, headerRowNumber)
        Dim columnType As Long() = getColumnType(sheet, columns, headerRowNumber)  
        Return generateColumn(sheet.Name, columnName, columnType)

    End Function

    Private Function generateColumn(sheetName As String, columnName As String(), columnType As Long()) As System.Data.DataTable
        Dim result As System.Data.DataTable = New System.Data.DataTable(sheetName)

        For i As Long = 0 To columnName.Count - 1
            Select Case columnType(i)
                Case 1
                    result.Columns.Add(columnName(i), GetType(String))
                Case 2
                    result.Columns.Add(columnName(i), GetType(String))
                Case 8
                    result.Columns.Add(columnName(i), GetType(String))
                Case 9
                    result.Columns.Add(columnName(i), GetType(String))
                Case 7
                    result.Columns.Add(columnName(i), GetType(String))
                Case 5
                    result.Columns.Add(columnName(i), GetType(String))
                Case 6
                    result.Columns.Add(columnName(i), GetType(String))
                Case 4
                    result.Columns.Add(columnName(i), GetType(String))

            End Select

        Next
        Return result
    End Function

    Private Function getColumnName(sheet As SpreadsheetGear.IWorksheet, columns As Long, headerRowNumber As Long) As String()
        Dim result As System.Collections.Generic.List(Of String) = New System.Collections.Generic.List(Of String)
        For i As Long = 0 To columns - 1
            If sheet.Cells(headerRowNumber - 1, i).Formula.Length <= 0 Then
                result.Add("column_" & i)
            Else
                result.Add(sheet.Cells(headerRowNumber - 1, i).Formula)
            End If
        Next
        Return result.ToArray
    End Function

    Private Function getColumnType(sheet As SpreadsheetGear.IWorksheet, columns As Long, headerRowNumber As Long) As Long()
        Dim result As System.Collections.Generic.List(Of Long) = New System.Collections.Generic.List(Of Long)
        For i As Long = 0 To columns - 1
            result.Add(sheet.Cells(headerRowNumber - 1, i).NumberFormatType)
        Next
        Return result.ToArray
    End Function


End Class
