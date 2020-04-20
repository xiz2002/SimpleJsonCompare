Attribute VB_Name = "TableModule"
''
' Json Simple Compare
' https://github.com/xiz2002/SimpleJsonCompare
'
' TableModule
'
' @class TableModule
' @author Lee Daho
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' MIT License
'
' Copyright (c) 2020 Lee Daho
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Private Const T_FIELD_COLUMN = "Field"
Private Const T_VALUE_COLUMN = "Value"

''
' Check Exists Table for tableName
'
' @method isSheet
' @param {Worksheet} sheet
' @param {String} tableName
' @return {Integer} TableIndex
''
Public Function isTable(ByRef sheet As Worksheet, ByVal tableName As String) As Integer
    Dim cnt As Integer
    ' 시트내 테이블 갯수 취득
    cnt = sheet.ListObjects.Count
       
    ' 테이블 갯수가 0건인 경우, e0 점프
    If Not CBool(cnt) Then GoTo e0
    
    ' 테이블 갯수 만큼, 루프
    Do
        ' 테이블 이름이 존재하면 루프종료
        If sheet.ListObjects(cnt).Name = tableName Then Exit Do
        cnt = cnt - 1
    Loop Until cnt = 0
    
    isTable = cnt
    Exit Function
e0:
    isTable = 0
End Function

''
' create Table on FileName
'
' @method createTable
' @param {Worksheet} sheet
' @param {String} tableName
' @param {Integer|Empty} xlTheme 테마 컬러
' @return {ListObject} Table
''
Public Function createTable(ByRef sheet As Worksheet, ByVal tableName As String, Optional xlTheme As Integer = 11) As ListObject
    Dim table As ListObject, tableIndex As Integer
    
    ' 테이블 Index 취득
    tableIndex = isTable(sheet, tableName)
    
    ' 테이블 Index가 존재하는경우 삭제
    If CBool(tableIndex) Then
        Debug.Print "Already Exists Table " + tableName
        sheet.ListObjects(tableIndex).Delete
    End If

    ' 테이블 생성
    Set table = sheet.ListObjects.Add(xlSrcRange, Range("$A$1"), , xlNo, , "TableStyleLight" + CStr(xlTheme))
    ' 테이블 이름 설정
    With table
        .Name = tableName
        ' 테이블 컬럼 추가 및, 컬럼 명 설정
        With .ListColumns
            .Add
            .item(1).Name = T_FIELD_COLUMN
            .item(2).Name = T_VALUE_COLUMN
         End With
    End With
    
    ' Return
    Set createTable = table
End Function

''
' create Table on FileName
'
' @method createTable
' @param {Worksheet} sheet 시트
' @param {String} queryName 출력할 쿼리 명
' @param {String} tableName 테이블출력명
' @param {Integer|Empty} xlTheme 테마 컬러
''
Public Sub createTableForQuery(ByRef sheet As Worksheet, ByVal queryName As String, Optional xlTheme As Integer = 11)
    Dim table As QueryTable, tableIndex As Integer
    
    ' 테이블 Index 취득
    tableIndex = isTable(sheet, queryName)
    
    ' 테이블 Index가 존재하는경우 삭제
    If CBool(tableIndex) Then
        Debug.Print "Already Exists Table " + queryName
        sheet.ListObjects(tableIndex).Delete
    End If
    
    ' 테이블 생성
    Set table = sheet.ListObjects.Add( _
        SourceType:=xlSrcExternal, _
        Source:="OLEDB; Provider=Microsoft.Mashup.OleDb.1; Data Source=$Workbook$; Location=" & queryName & "; Extended Properties=""""", _
        Destination:=Range("$A$1"), _
        TableStyleName:="TableStyleLight" + CStr(xlTheme)).QueryTable
    
    ' 테이블 설정
    With table
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & queryName & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = queryName
        .Refresh BackgroundQuery:=False
    End With
End Sub

''
' Add Data to Table
'
' @method addDataToTable
' @param {ListObject} tbl 테이블
' @param {Array} data 테이블에 입력할 데이터
''
Public Sub addDataToTable(ByRef tbl As ListObject, ByVal data As Variant)
    Dim lastRange As Excel.Range, index As Long
    ' 테이블에 행 추가 및, 마지막 범위 취득
    With tbl
        .ListRows.Add
        Set lastRange = .DataBodyRange.Rows(.ListRows.Count)
    End With
    
    ' 마지막 범위의 컬럼 갯수 만큼 루프
    For index = 0 To lastRange.Columns.Count - 1
        ' 데이터 입력
        lastRange.Columns.Cells.item(1, (index + 1)).Value = data(index)
    Next index
End Sub

''
' Add Data to Table (For Table Name)
'
' @method addDataToTableForName
' @param {String} tblName 테이블 명
' @param {Array} data 테이블에 입력할 데이터
' @param {String|Null} shName 시트 명
''
Public Sub addDataToTableForName(ByVal tblName As String, ByVal data As Variant, Optional ByVal shName As String = "")
    Dim tbl As ListObject
    Dim lastRange As Excel.Range, index As Long
    
    If Trim(shName & vbNullString) = vbNullString Then
        ' 시트명이 존재하지 않으면, 현재 시트에서 테이블 취득
        Set tbl = ActiveSheet.ListObjects(tblName)
    Else
        ' 시트명이 존재하면, 특정 시트에서, 테이블 취득
        Set tbl = ActiveWorkbook.Worksheets(shName).ListObjects(tblName)
    End If
   
    ' 테이블에 행 추가 및, 마지막 범위 취득
    With tbl
        .ListRows.Add
        Set lastRange = .DataBodyRange.Rows(.ListRows.Count)
    End With
    
    ' 마지막 범위의 컬럼 갯수 만큼 루프
    For index = 0 To lastRange.Columns.Count - 1
        ' 데이터 입력
        lastRange.Columns.Cells.item(1, (index + 1)).Value = data(index)
    Next index
End Sub
