Attribute VB_Name = "QueryModule"
''
' Json Simple Compare
' https://github.com/xiz2002/SimpleJsonCompare
'
' QueryModule
'
' @class QueryModule
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

Public Const PREFIX_TRANSPOSE = "TP_"
Public Const PREFIX_COMBINE = "CB_"

''
' Check Exists Query
'
' @method isQuery
' @param {String} queryName
' @return {Boolean} False: �̻�����, True: ��������
''
Public Function isQuery(ByVal queryName As String) As Boolean
    On Error GoTo e0
    ' �����ϴ��� Ȯ��
    If CBool(Len(ActiveWorkbook.Queries(queryName).Name)) Then
        ' �������� True
        isQuery = True
    End If
    Exit Function
e0:
    ' �̻����� False
    isQuery = False
End Function

''
' Remove Query
'
' @method removeSheet
' @param {String} ��Ʈ��
' @return {Integer} -1: ������ / 0: �������� / 1: �̻�����
''
Public Function removeQuery(ByVal queryName As String) As Integer
    On Error GoTo e0
    
    ' ���� ���� ���� üũ
    If isQuery(queryName) Then
        Call ActiveWorkbook.Queries(queryName).Delete
        ' �������� 0
        removeQuery = 0
        Exit Function
    End If
    
    ' ������ -1
    removeQuery = -1
    Exit Function
e0:
    ' �̻����� 1
    removeQuery = 1
End Function

''
' Create Query
'
' @method createQuery
' @param {String} tableName ���̺� ��
''
Public Sub createQuery(ByVal tableName As String)
    Dim sFormula As String
    ' ���� �ۼ�
    sFormula = _
        "let Source = Excel.CurrentWorkbook(){[Name=""" + tableName + """]}[Content] in Source"
    ' ������ �̹� �����ϴ� ��� ����
    If isQuery(tableName) Then
        Call ActiveWorkbook.Queries(tableName).Delete
    End If
    ' �ۼ��� ���� �߰�
    Call ActiveWorkbook.Queries.Add(tableName, sFormula)
End Sub

''
' Create Transpose Query
'
' ���̺��� �࿭�� ��ȯ�� ������ �ۼ��Ѵ�.
'
' @method transposeQuery
' @param tableName ���̺� ��
''
Public Sub createTransposeQuery(ByVal tableName As String)
    Dim isExistQuery As Boolean, sFormula As String
    ' Query ���翩��
    isExistQuery = isQuery(tableName)
    
    If Not isExistQuery Then
        ' ������ �̹� �����ϴ°�� ����
        sFormula = "let Source = tableName + "","" "
    Else
        ' ������ �������� �ʴ°�� ����
        sFormula = "let Source = Excel.CurrentWorkbook(){[Name=""" + tableName + """]}[Content],"
    End If
    
    ' ���� �ۼ�
    sFormula = sFormula _
        & Chr(13) + Chr(10) + Chr(9) & "Transpose = Table.Transpose(Source)," _
        & Chr(13) + Chr(10) + Chr(9) & "Result = Table.PromoteHeaders(Transpose, [PromoteAllScalars=true])" _
        & Chr(13) + Chr(10) & "in Result"
        
    ' Transpose������ �̹� �����ϴ� ��� ����
    If isQuery(PREFIX_TRANSPOSE + tableName) Then
        Call ActiveWorkbook.Queries(PREFIX_TRANSPOSE + tableName).Delete
    End If
        
    ' �ۼ��� ���� �߰�
    Call ActiveWorkbook.Queries.Add(PREFIX_TRANSPOSE + tableName, sFormula)
End Sub

''
' Create Combine Qeury Table
'
' @method createCombineQuery
' @param {String} QueryName ������ ���� �̸�
' @param {Array} tableNames ������ ���̺� �̸� ����Ʈ
''
Public Sub createCombineQuery(ByVal queryName As String, ByRef tableNames As Variant)
    Dim sFormula As String, index As Long
    
    ' ���̺� �̸� ����Ʈ�� ���� ���� �ʴ� ��� ó�� ����
    If ((LBound(tableNames) = 0) And (UBound(tableNames) = -1)) Then
       Debug.Print "Array was not provided"
       Exit Sub
    End If
   
    ' ���̺� �̸� ����Ʈ�� ù��° ��Ұ� ���� ���� �ʴ°�� ó�� ����
    If IsEmpty(tableNames(0)) Then
       Debug.Print "'Nothing' was passed in"
       Exit Sub
    End If

    ' ��ĥ ���̺� ��ȯ
    For index = LBound(tableNames) To UBound(tableNames)
        ' ������ ���� ���� �ʴ°�� ����
        If Not isQuery(tableNames(index)) Then
            Debug.Print "Not Exist Query:" + tableNames(index)
            Exit For
        End If
        ' Transpose���� ����
        Call createTransposeQuery(tableNames(index))
    Next index

    ' Combine������ �̹� �����ϴ� ��� ����
    If isQuery(PREFIX_COMBINE + queryName) Then
        Call ActiveWorkbook.Queries(PREFIX_COMBINE + queryName).Delete
    End If
    
    ' Combine���� �ۼ�
    sFormula = "let Source = Table.Combine({"
    
    ' ��ĥ ���̺� �߰�
    For index = LBound(tableNames) To UBound(tableNames)
'        ' Transpose������ ���� ���� �ʴ°�� ����
'        If Not isQuery(PREFIX_TRANSPOSE + tableNames(index)) Then
'            Debug.Print "Not Exist TransposeQuery:" + PREFIX_TRANSPOSE + tableNames(index)
'            Exit For
'        End If
        
        ' ���̺� �� �߰�
        sFormula = sFormula + PREFIX_TRANSPOSE + tableNames(index) + ","
        
        ' ������ �ε����� ��� "," ����
        If index = UBound(tableNames) Then
            sFormula = Mid(sFormula, 1, InStrRev(sFormula, ",") - 1) + "}),"
        End If
    Next index

    ' ����� ������ ����
    sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "DemoteHeader = Table.DemoteHeaders(Source),"
    ' �࿭ ����
    sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "Transpose = Table.Transpose(DemoteHeader),"
    ' �÷��� ����
    sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "RenameColumns = Table.RenameColumns(Transpose,{"
    
    ' ù��° �÷� ����(Fiexd Field)
    sFormula = sFormula & "{" & """Column1""," & """Field_Name""" & "},"
    
    ' ������ �÷��� ����Ʈ
    For index = LBound(tableNames) To UBound(tableNames)
'        ' Transpose������ ���� ���� �ʴ°�� ����
'        If Not isQuery(PREFIX_TRANSPOSE + tableNames(index)) Then
'            Debug.Print "Not Exist Column:" + PREFIX_TRANSPOSE + tableNames(index)
'            Exit For
'        End If
        
        ' �÷� �� ����
        sFormula = sFormula & "{" & """Column" & index + 2 & """," & """" & tableNames(index) & ".Value""" & "},"
        
        ' ������ �ε����� ��� "," ����
        If index = UBound(tableNames) Then
            sFormula = Mid(sFormula, 1, InStrRev(sFormula, ",") - 1) + "}),"
        End If
    Next index
    
    ' �� ��� �÷� �߰�
    If 1 < UBound(tableNames) Then
        ' 3���� ���� ���� ���, 1.Value eq 3.Value, 2.Value eq 3.Value �� ���� �߰�
        sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "addColumn3 = Table.AddColumn(RenameColumns, ""isDiff(1 eq 3)"", each Value.Compare([File_1.Value], [File_3.Value])),"
        sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "addColumn2 = Table.AddColumn(addColumn3, ""isDiff(2 eq 3)"", each Value.Compare([File_2.Value], [File_3.Value])),"
        sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "addColumn1 = Table.AddColumn(addColumn2, ""isDiff(1 eq 2)"", each Value.Compare([File_1.Value], [File_2.Value])),"
        sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "Results = Table.TransformColumnTypes(addColumn1, {" _
        & "{""isDiff(1 eq 3)"", type logical}," & "{""isDiff(2 eq 3)"", type logical}," & "{""isDiff(1 eq 2)"", type logical}" & "})"
    Else
        sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "addColumn1 = Table.AddColumn(RenameColumns, ""isDiff(1 eq 2)"", each Value.Compare([File_1.Value], [File_2.Value])),"
        sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "Results = Table.TransformColumnTypes(addColumn1, {{""isDiff(1 eq 2)"", type logical}})"
    End If

    ' ���
    sFormula = sFormula & Chr(13) + Chr(10) & "in Results"
    
    ' Query ����
    Call ActiveWorkbook.Queries.Add(PREFIX_COMBINE + queryName, sFormula)
    
    ' Query ���
    Call queryToWsTbl("Result_Compare", queryName)
End Sub

''
' Create WorkSheet Table For Query
'
' @method
' @param sheetName ��Ʈ��
' @param queryName ����� ���� ��
''
Private Sub queryToWsTbl(ByVal sheetName As String, ByVal queryName As String)
    Dim sheet As Worksheet
    ' Sheet �߰�
    Set sheet = SheetModule.createSheet(sheetName, xlThemeColorAccent6)
    ' Table �߰�
    Call createTableForQuery(sheet, PREFIX_COMBINE + queryName)
    ' deAllocate
    Set sheet = Nothing
End Sub
