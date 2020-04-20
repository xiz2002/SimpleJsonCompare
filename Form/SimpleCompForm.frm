VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SimpleCompForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13365
   OleObjectBlob   =   "SimpleCompForm.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "SimpleCompForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' Json Simple Compare
' https://github.com/xiz2002/SimpleJsonCompare
'
' Json Simple Compare Form
'
' @ dependence
' - Module: QueryModule.bas, FileModule.bas
' - ClassModule: JsonToTable.cls
'
' @class SimpleCompForm
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
Private m_JsonTableList As Collection                                                   ' JsonToTable Object Collection

'' Constructor''
Private Sub UserForm_Initialize()
    ' 콜렉션 생성
    Set m_JsonTableList = New Collection
    ' 데이터 수동계산 설정
    Application.Calculation = xlManual
End Sub

'' Destory ''
Private Sub UserForm_Terminate()
    On Error Resume Next
    ' 콜렉션 내부 객체 해제
    Call collectionCleaner
    ' 콜렉션 해제
    Set m_JsonTableList = Nothing
    ' 데이터 자동계산 설정
    Application.Calculation = xlAutomatic
End Sub

''
' Collection Cleaner
' 콜렉션 내부 오브젝트 삭제처리
'
' @method collectionCleaner
''
Private Sub collectionCleaner()
    Dim index As Integer
    
    For index = 1 To m_JsonTableList.Count
        m_JsonTableList.Remove 1
    Next
End Sub

''
' Execute Button Event
' 실행버튼 이벤트
'
' @method btnExecute_Click
''
Private Sub btnExecute_Click()
    Dim ctrl As Control, instance As JsonToTable
    Dim ctrlTag As String, jsonData As String
    Dim tableNames As Variant, index As Long
    
    ' Clean for Collection
    Call collectionCleaner
    
    ' 폼 컨트롤 검색
    For Each ctrl In Me.Controls
        ' 텍스트 박스인 경우, 텍스트 내용 검증 후, 객체 생성
        If TypeOf ctrl Is MSForms.TextBox Then
            ' 태그 취득
            ctrlTag = ctrl.Tag
        
            If Trim(ctrl.Text & vbNullString) = vbNullString And Not ctrl.Name = "tbFilePath3" Then
                ' File_3 이외 파일 경로가 존재 하지 않는경우, e0으로 점프
                GoTo e0
            ElseIf Trim(ctrl.Text & vbNullString) = vbNullString And ctrl.Name = "tbFilePath3" Then
                ' File_3 파일 경로가 존재 하지 않는경우
                ' 쿼리 삭제
                Call QueryModule.removeQuery(ctrlTag)
                ' Transpose쿼리 삭제
                Call QueryModule.removeQuery(QueryModule.PREFIX_TRANSPOSE + ctrlTag)
                ' 시트 삭제
                Call SheetModule.removeSheet(ctrlTag)
                ' 반복문 종료
                Exit For
            End If
            
            ' 파일경로를 이용해, Json 파일 내용 취득
            jsonData = FileModule.readText(ctrl.Text)
            ' fileName = Mid(Dir(ctrl.Text), 1, InStrRev(LCase(Dir(ctrl.Text)), ".") - 1)
            
            ' Json파일 내용이 존재 하지 않는경우, e1으로 점프
            If Trim(jsonData & vbNullString) = vbNullString Then
                GoTo e1
            End If
            
            ' JsonToTable 클래스 생성 및 초기화
            Set instance = New JsonToTable
            instance.jsonData = jsonData
            instance.tableName = ctrlTag
            ' 인스턴스를 콜렉션에 추가
            m_JsonTableList.Add item:=instance, key:=ctrlTag
            ' 메모리 해제
            Set instance = Nothing
            
            If IsEmpty(tableNames) Then
                ' 테이블 이름 배열 생성
                tableNames = Array(ctrlTag)
            ElseIf UBound(tableNames) >= 0 Then
                ' 테이블 이름 배열이 존재하면, 배열 사이즈 조정 및, 테이블 이름 추가
                index = UBound(tableNames) + 1
                ReDim Preserve tableNames(index)
                tableNames(index) = ctrlTag
            End If

        End If
    Next
    
    ' Json 파일 내용을 테이블로 생성
    For Each instance In m_JsonTableList
        Call instance.ConvertJsonToTable
    Next
    
    ' Query 생성
    For Each instance In m_JsonTableList
        Call QueryModule.createQuery(instance.tableName)
    Next
    
    ' Qeury Combine
    Call QueryModule.createCombineQuery("Compare", tableNames)
    
    ' 데이터 자동계산 설정
    Application.Calculation = xlAutomatic
    Exit Sub
e0:
    lbInfo.Caption = "Error: " + ctrlTag + " Path Empty."
    Application.Calculation = xlAutomatic
    Exit Sub
e1:
    lbInfo.Caption = "Error: " + ctrlTag + " File Wrong."
    Application.Calculation = xlAutomatic
End Sub

''
' Dialog Open get File Path For File 1
'
' @method btnOpenDialog1_Click
''
Private Sub btnOpenDialog1_Click()
    tbFilePath1.Value = FileModule.openDialog()
End Sub

''
' Dialog Open get File Path For File 2
'
' @method btnOpenDialog2_Click
''
Private Sub btnOpenDialog2_Click()
    tbFilePath2.Value = FileModule.openDialog()
End Sub

''
' Dialog Open get File Path For File 3
'
' @method btnOpenDialog3_Click
''
Private Sub btnOpenDialog3_Click()
    tbFilePath3.Value = FileModule.openDialog()
End Sub
