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
    ' ��Ʈ�� ���̺� ���� ���
    cnt = sheet.ListObjects.Count
       
    ' ���̺� ������ 0���� ���, e0 ����
    If Not CBool(cnt) Then GoTo e0
    
    ' ���̺� ���� ��ŭ, ����
    Do
        ' ���̺� �̸��� �����ϸ� ��������
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
' @param {Integer|Empty} xlTheme �׸� �÷�
' @return {ListObject} Table
''
Public Function createTable(ByRef sheet As Worksheet, ByVal tableName As String, Optional xlTheme As Integer = 11) As ListObject
    Dim table As ListObject, tableIndex As Integer
    
    ' ���̺� Index ���
    tableIndex = isTable(sheet, tableName)
    
    ' ���̺� Index�� �����ϴ°�� ����
    If CBool(tableIndex) Then
        Debug.Print "Already Exists Table " + tableName
        sheet.ListObjects(tableIndex).Delete
    End If

    ' ���̺� ����
    Set table = sheet.ListObjects.Add(xlSrcRange, Range("$A$1"), , xlNo, , "TableStyleLight" + CStr(xlTheme))
    ' ���̺� �̸� ����
    With table
        .Name = tableName
        ' ���̺� �÷� �߰� ��, �÷� �� ����
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
' @param {Worksheet} sheet ��Ʈ
' @param {String} queryName ����� ���� ��
' @param {String} tableName ���̺���¸�
' @param {Integer|Empty} xlTheme �׸� �÷�
''
Public Sub createTableForQuery(ByRef sheet As Worksheet, ByVal queryName As String, Optional xlTheme As Integer = 11)
    Dim table As QueryTable, tableIndex As Integer
    
    ' ���̺� Index ���
    tableIndex = isTable(sheet, queryName)
    
    ' ���̺� Index�� �����ϴ°�� ����
    If CBool(tableIndex) Then
        Debug.Print "Already Exists Table " + queryName
        sheet.ListObjects(tableIndex).Delete
    End If
    
    ' ���̺� ����
    Set table = sheet.ListObjects.Add( _
        SourceType:=xlSrcExternal, _
        Source:="OLEDB; Provider=Microsoft.Mashup.OleDb.1; Data Source=$Workbook$; Location=" & queryName & "; Extended Properties=""""", _
        Destination:=Range("$A$1"), _
        TableStyleName:="TableStyleLight" + CStr(xlTheme)).QueryTable
    
    ' ���̺� ����
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
' @param {ListObject} tbl ���̺�
' @param {Array} data ���̺� �Է��� ������
''
Public Sub addDataToTable(ByRef tbl As ListObject, ByVal data As Variant)
    Dim lastRange As Excel.Range, index As Long
    ' ���̺� �� �߰� ��, ������ ���� ���
    With tbl
        .ListRows.Add
        Set lastRange = .DataBodyRange.Rows(.ListRows.Count)
    End With
    
    ' ������ ������ �÷� ���� ��ŭ ����
    For index = 0 To lastRange.Columns.Count - 1
        ' ������ �Է�
        lastRange.Columns.Cells.item(1, (index + 1)).Value = data(index)
    Next index
End Sub

''
' Add Data to Table (For Table Name)
'
' @method addDataToTableForName
' @param {String} tblName ���̺� ��
' @param {Array} data ���̺� �Է��� ������
' @param {String|Null} shName ��Ʈ ��
''
Public Sub addDataToTableForName(ByVal tblName As String, ByVal data As Variant, Optional ByVal shName As String = "")
    Dim tbl As ListObject
    Dim lastRange As Excel.Range, index As Long
    
    If Trim(shName & vbNullString) = vbNullString Then
        ' ��Ʈ���� �������� ������, ���� ��Ʈ���� ���̺� ���
        Set tbl = ActiveSheet.ListObjects(tblName)
    Else
        ' ��Ʈ���� �����ϸ�, Ư�� ��Ʈ����, ���̺� ���
        Set tbl = ActiveWorkbook.Worksheets(shName).ListObjects(tblName)
    End If
   
    ' ���̺� �� �߰� ��, ������ ���� ���
    With tbl
        .ListRows.Add
        Set lastRange = .DataBodyRange.Rows(.ListRows.Count)
    End With
    
    ' ������ ������ �÷� ���� ��ŭ ����
    For index = 0 To lastRange.Columns.Count - 1
        ' ������ �Է�
        lastRange.Columns.Cells.item(1, (index + 1)).Value = data(index)
    Next index
End Sub
