VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonToTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Json Simple Compare
' https://github.com/xiz2002/SimpleJsonCompare
'
' JsonToTable
'
' @class JsonToTable
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

' 전역변수
Private m_JsonData As String                                    ' Json 파일 내용
Private m_TableName As String                                   ' Table 명
Private m_Table As ListObject                                   ' Table 객체
Private m_Data As Object                                        ' Json 파일 내용 해석 결과 (Dictionary or Collection)

''
' JsonData(RAW) Setter
' @param {String} data JsonRawData
''
Property Let jsonData(data As String)
    m_JsonData = data
End Property

''
' Getter TableName
' @param {String} tableName 테이블 명
''
Property Get tableName() As String
    tableName = m_TableName
End Property

''
' Setter TableName
' @param {String} tableName 테이블 명
''
Property Let tableName(tableName As String)
    m_TableName = tableName
End Property

'' 생성자''
Private Sub Class_Initialize()
End Sub

'' 파괴자''
Private Sub Class_Terminate()
    Set m_Table = Nothing
    Set m_Data = Nothing
End Sub

''
' Json Data To Table
'
' @method ConvertJsonToTable
' @param {String|Empty} Json파일 내용
' @param {String|Empty} Table 이름
''
Public Sub ConvertJsonToTable(Optional ByVal jsonData = "", Optional ByVal tableName = "")
    ' 인수가 존재하는 경우, Json 파일 내용 설정
    If Not Trim(jsonData & vbNullString) = vbNullString Then
        m_JsonData = jsonData
    End If
    
    ' 인수가 존재하는 경우, 테이블 이름 설정
    If Not Trim(tableName & vbNullString) = vbNullString Then
        m_TableName = tableName
    End If
    
    ' 파일 내용 또는 테이블 이름이 존재하지 않는경우 처리 종료
     If Trim(m_JsonData & vbNullString) = vbNullString Or Trim(m_TableName & vbNullString) = vbNullString Then
        Exit Sub
    End If
        
    ' 테이블 작성
    Set m_Table = TableModule.createTable(SheetModule.createSheet(m_TableName), m_TableName)
    
    ' Json 파일 내용 취득
    Set m_Data = JsonConverter.ParseJson(m_JsonData)
    
    ' Json 파일 내용 분해
    Call parsedDataDivision(m_Data)
End Sub

''
' Json Data Type Division
'
' @method parsedDataDivision
' @param {Dictionary|Collection} obj JsonParsing데이터
' @param {String|Empty} pField 상위필드명
''
Private Sub parsedDataDivision(ByVal obj As Object, Optional ByVal pField As String = "")
    If VBA.TypeName(obj) = "Dictionary" Then
        ' 자료타입이 사전인경우
        Dim key As Variant
        
        ' Key 길이 만큼 루프
        For Each key In obj.Keys
            Call divDataAndInsert(obj(key), _
                                IIf(Trim(pField & vbNullString) = vbNullString, key, pField & "." & key))
        Next
    ElseIf VBA.TypeName(obj) = "Collection" Then
        ' 자료타입이 콜렉션인경우
        Dim index As Integer
        
        ' 인덱스 수 만큼 루프
        For index = 1 To obj.Count
            Call divDataAndInsert(obj(index), _
                                IIf(Trim(pField & vbNullString) = vbNullString, CStr(index), pField & "." & CStr(index)))
        Next index
    Else
        ' 그외 자료타입
        Debug.Print "Type Miss:" + pField + CStr(obj)
    End If
End Sub

''
' Json Item Division And Write Table
'
' @method divDataAndInsert
' @param {Variant} data 값
' @param {String} field 필드
''
Private Function divDataAndInsert(ByVal data As Variant, ByVal field As String)
    Select Case VBA.VarType(data)
    Case VBA.vbNull
        ' 자료타입이 Null인경우
        ' Debug.Print "[Field]:" + CStr(field) + " [Value]:" + "Null" + " [Type]: Null"
        Call TableModule.addDataToTable(m_Table, Array(CStr(field), Null))
    Case VBA.vbDate
        ' 자료타입이 Date인경우
        ' Debug.Print "[Field]:" + CStr(field) + " [Value]:" + CStr(data) + " [Type]: Date"
        Call TableModule.addDataToTable(m_Table, Array(CStr(field), data))
    Case VBA.vbString
        ' 자료타입이 문자열인경우
        ' Debug.Print "[Field]:" + CStr(field) + " [Value]:" + CStr(data) + " [Type]: String"
        Call TableModule.addDataToTable(m_Table, Array(CStr(field), CStr("'" + data)))
    Case VBA.vbBoolean
        ' 자료타입이 Bool인경우
        ' Debug.Print "[Field]:" + CStr(field) + " [Value]:" + CStr(data) + " [Type]: Boolean"
        Call TableModule.addDataToTable(m_Table, Array(CStr(field), data))
    Case VBA.vbArray
        ' 자료타입이 배열인경우
        ' Debug.Print "[Field]:" + CStr(field) + " [Value]:" + CStr(data) + " [Type]: Array"
        Call TableModule.addDataToTable(m_Table, Array(CStr(field), data))
    Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
        ' 자료타입이 숫자인경우
        ' Debug.Print "[Field]:" + CStr(field) + " [Value]:" + CStr(data) + " [Type]: Number"
        Call TableModule.addDataToTable(m_Table, Array(CStr(field), data))
    Case VBA.vbObject
        If VBA.TypeName(data) = "Dictionary" Then
            ' 자료타입이 사전인경우
            
            ' Empty Check
            If Not CBool(data.Count) Then
                ' Debug.Print "[Field]:" + CStr(field) + " [Value]:" + "Empty" + " [Type]: Dictionary"
                Call TableModule.addDataToTable(m_Table, Array(CStr(field), Empty))
                Exit Function
            End If
            
            Call parsedDataDivision(data, field)
        ElseIf VBA.TypeName(data) = "Collection" Then
            ' 자료타입이 콜렉션인경우
            
            ' Empty Check
            If Not CBool(data.Count) Then
                ' Debug.Print "[Field]:" + CStr(field) + " [Value]:" + "Empty" + " [Type]: Collection"
                Call TableModule.addDataToTable(m_Table, Array(CStr(field), Empty))
                Exit Function
            End If
            
            ' Note: 굳이 리스트 를 추출 할 필요가 있는가?
            If VBA.VarType(data(1)) = VBA.vbObject Then
                ' 콜렉션 내부 자료형이 오브젝트인 경우
                Call parsedDataDivision(data, field)
            Else
                ' 콜렉션 내부 자료형이 오브젝트 이외인 경우
                Dim accItem As String, item As Variant
                ' 내부 자료 취득 및 문자열로 변경
                For Each item In data
                    accItem = IIf(Trim(accItem & vbNullString) = vbNullString, CStr(item), accItem & "," & CStr(item))
                Next
                
                Call TableModule.addDataToTable(m_Table, Array(CStr(field), accItem))
            End If
        End If
    End Select
End Function
