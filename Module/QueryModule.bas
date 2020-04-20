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
' @return {Boolean} False: 이상종료, True: 정상종료
''
Public Function isQuery(ByVal queryName As String) As Boolean
    On Error GoTo e0
    ' 존재하는지 확인
    If CBool(Len(ActiveWorkbook.Queries(queryName).Name)) Then
        ' 정상종료 True
        isQuery = True
    End If
    Exit Function
e0:
    ' 이상종료 False
    isQuery = False
End Function

''
' Remove Query
'
' @method removeSheet
' @param {String} 시트명
' @return {Integer} -1: 미존재 / 0: 정상종료 / 1: 이상종료
''
Public Function removeQuery(ByVal queryName As String) As Integer
    On Error GoTo e0
    
    ' 쿼리 존재 여부 체크
    If isQuery(queryName) Then
        Call ActiveWorkbook.Queries(queryName).Delete
        ' 정상종료 0
        removeQuery = 0
        Exit Function
    End If
    
    ' 미존재 -1
    removeQuery = -1
    Exit Function
e0:
    ' 이상종료 1
    removeQuery = 1
End Function

''
' Create Query
'
' @method createQuery
' @param {String} tableName 테이블 명
''
Public Sub createQuery(ByVal tableName As String)
    Dim sFormula As String
    ' 쿼리 작성
    sFormula = _
        "let Source = Excel.CurrentWorkbook(){[Name=""" + tableName + """]}[Content] in Source"
    ' 쿼리가 이미 존재하는 경우 삭제
    If isQuery(tableName) Then
        Call ActiveWorkbook.Queries(tableName).Delete
    End If
    ' 작성된 쿼리 추가
    Call ActiveWorkbook.Queries.Add(tableName, sFormula)
End Sub

''
' Create Transpose Query
'
' 테이블의 행열이 변환된 쿼리를 작성한다.
'
' @method transposeQuery
' @param tableName 테이블 명
''
Public Sub createTransposeQuery(ByVal tableName As String)
    Dim isExistQuery As Boolean, sFormula As String
    ' Query 존재여부
    isExistQuery = isQuery(tableName)
    
    If Not isExistQuery Then
        ' 쿼리가 이미 존재하는경우 참조
        sFormula = "let Source = tableName + "","" "
    Else
        ' 쿼리가 존재하지 않는경우 생성
        sFormula = "let Source = Excel.CurrentWorkbook(){[Name=""" + tableName + """]}[Content],"
    End If
    
    ' 쿼리 작성
    sFormula = sFormula _
        & Chr(13) + Chr(10) + Chr(9) & "Transpose = Table.Transpose(Source)," _
        & Chr(13) + Chr(10) + Chr(9) & "Result = Table.PromoteHeaders(Transpose, [PromoteAllScalars=true])" _
        & Chr(13) + Chr(10) & "in Result"
        
    ' Transpose쿼리가 이미 존재하는 경우 삭제
    If isQuery(PREFIX_TRANSPOSE + tableName) Then
        Call ActiveWorkbook.Queries(PREFIX_TRANSPOSE + tableName).Delete
    End If
        
    ' 작성된 쿼리 추가
    Call ActiveWorkbook.Queries.Add(PREFIX_TRANSPOSE + tableName, sFormula)
End Sub

''
' Create Combine Qeury Table
'
' @method createCombineQuery
' @param {String} QueryName 생성할 쿼리 이름
' @param {Array} tableNames 합쳐질 테이블 이름 리스트
''
Public Sub createCombineQuery(ByVal queryName As String, ByRef tableNames As Variant)
    Dim sFormula As String, index As Long
    
    ' 테이블 이름 리스트가 존재 하지 않는 경우 처리 종료
    If ((LBound(tableNames) = 0) And (UBound(tableNames) = -1)) Then
       Debug.Print "Array was not provided"
       Exit Sub
    End If
   
    ' 테이블 이름 리스트의 첫번째 요소가 존재 하지 않는경우 처리 종료
    If IsEmpty(tableNames(0)) Then
       Debug.Print "'Nothing' was passed in"
       Exit Sub
    End If

    ' 합칠 테이블 변환
    For index = LBound(tableNames) To UBound(tableNames)
        ' 쿼리가 존재 하지 않는경우 제외
        If Not isQuery(tableNames(index)) Then
            Debug.Print "Not Exist Query:" + tableNames(index)
            Exit For
        End If
        ' Transpose쿼리 생성
        Call createTransposeQuery(tableNames(index))
    Next index

    ' Combine쿼리가 이미 존재하는 경우 삭제
    If isQuery(PREFIX_COMBINE + queryName) Then
        Call ActiveWorkbook.Queries(PREFIX_COMBINE + queryName).Delete
    End If
    
    ' Combine쿼리 작성
    sFormula = "let Source = Table.Combine({"
    
    ' 합칠 테이블 추가
    For index = LBound(tableNames) To UBound(tableNames)
'        ' Transpose쿼리가 존재 하지 않는경우 제외
'        If Not isQuery(PREFIX_TRANSPOSE + tableNames(index)) Then
'            Debug.Print "Not Exist TransposeQuery:" + PREFIX_TRANSPOSE + tableNames(index)
'            Exit For
'        End If
        
        ' 테이블 명 추가
        sFormula = sFormula + PREFIX_TRANSPOSE + tableNames(index) + ","
        
        ' 마지막 인덱스인 경우 "," 제외
        If index = UBound(tableNames) Then
            sFormula = Mid(sFormula, 1, InStrRev(sFormula, ",") - 1) + "}),"
        End If
    Next index

    ' 헤더를 행으로 변경
    sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "DemoteHeader = Table.DemoteHeaders(Source),"
    ' 행열 변경
    sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "Transpose = Table.Transpose(DemoteHeader),"
    ' 컬럼명 변경
    sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "RenameColumns = Table.RenameColumns(Transpose,{"
    
    ' 첫번째 컬럼 변경(Fiexd Field)
    sFormula = sFormula & "{" & """Column1""," & """Field_Name""" & "},"
    
    ' 변경할 컬럼명 리스트
    For index = LBound(tableNames) To UBound(tableNames)
'        ' Transpose쿼리가 존재 하지 않는경우 제외
'        If Not isQuery(PREFIX_TRANSPOSE + tableNames(index)) Then
'            Debug.Print "Not Exist Column:" + PREFIX_TRANSPOSE + tableNames(index)
'            Exit For
'        End If
        
        ' 컬럼 명 변경
        sFormula = sFormula & "{" & """Column" & index + 2 & """," & """" & tableNames(index) & ".Value""" & "},"
        
        ' 마지막 인덱스인 경우 "," 제외
        If index = UBound(tableNames) Then
            sFormula = Mid(sFormula, 1, InStrRev(sFormula, ",") - 1) + "}),"
        End If
    Next index
    
    ' 비교 결과 컬럼 추가
    If 1 < UBound(tableNames) Then
        ' 3개의 쿼리 비교의 경우, 1.Value eq 3.Value, 2.Value eq 3.Value 를 먼저 추가
        sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "addColumn3 = Table.AddColumn(RenameColumns, ""isDiff(1 eq 3)"", each Value.Compare([File_1.Value], [File_3.Value])),"
        sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "addColumn2 = Table.AddColumn(addColumn3, ""isDiff(2 eq 3)"", each Value.Compare([File_2.Value], [File_3.Value])),"
        sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "addColumn1 = Table.AddColumn(addColumn2, ""isDiff(1 eq 2)"", each Value.Compare([File_1.Value], [File_2.Value])),"
        sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "Results = Table.TransformColumnTypes(addColumn1, {" _
        & "{""isDiff(1 eq 3)"", type logical}," & "{""isDiff(2 eq 3)"", type logical}," & "{""isDiff(1 eq 2)"", type logical}" & "})"
    Else
        sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "addColumn1 = Table.AddColumn(RenameColumns, ""isDiff(1 eq 2)"", each Value.Compare([File_1.Value], [File_2.Value])),"
        sFormula = sFormula & Chr(13) + Chr(10) + Chr(9) & "Results = Table.TransformColumnTypes(addColumn1, {{""isDiff(1 eq 2)"", type logical}})"
    End If

    ' 결과
    sFormula = sFormula & Chr(13) + Chr(10) & "in Results"
    
    ' Query 생성
    Call ActiveWorkbook.Queries.Add(PREFIX_COMBINE + queryName, sFormula)
    
    ' Query 출력
    Call queryToWsTbl("Result_Compare", queryName)
End Sub

''
' Create WorkSheet Table For Query
'
' @method
' @param sheetName 시트명
' @param queryName 출력할 쿼리 명
''
Private Sub queryToWsTbl(ByVal sheetName As String, ByVal queryName As String)
    Dim sheet As Worksheet
    ' Sheet 추가
    Set sheet = SheetModule.createSheet(sheetName, xlThemeColorAccent6)
    ' Table 추가
    Call createTableForQuery(sheet, PREFIX_COMBINE + queryName)
    ' deAllocate
    Set sheet = Nothing
End Sub
