Attribute VB_Name = "SheetModule"
''
' Json Simple Compare
' https://github.com/xiz2002/SimpleJsonCompare
'
' SheetModule
'
' @class SheetModule
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

''
' Check Exists Sheet for Name
'
' @method isSheet
' @param {String} 시트명
' @return {Boolean} 시트 존재 여부
''
Public Function isSheet(ByVal sheetName As String) As Boolean
    ' 에러 발생시 점프할 위치 설정
    On Error GoTo e0
    ' 시트명 취득 후, 시트명이 존재하는 경우, True
    If CBool(Len(Sheets(sheetName).Name)) Then isSheet = True
    Exit Function
    
e0:
    isSheet = False
End Function

''
' Remove Sheet For Name
'
' @method removeSheet
' @param {String} 시트명
' @return {Integer} -1: 미존재 / 0: 정상종료 / 1: 이상종료
''
Public Function removeSheet(ByVal sheetName As String) As Integer
    Dim wbs As Worksheet
    ' 다이얼로그 비표시
    Application.DisplayAlerts = False
    
    ' 에러 발생시 점프 위치
    On Error GoTo e0
    
    ' 시트명 검색
    For Each wbs In Sheets
        With wbs
            If .Name = sheetName Then
                .Delete
                ' 삭제한 경우
                removeSheet = 0
                GoTo s0
            End If
        End With
    Next
    
    ' 삭제할 대상이 없는경우
    removeSheet = -1
s0:
    ' 정상 종료
    Application.DisplayAlerts = True
    Exit Function
e0:
    ' 이상 종료
    removeSheet = 1
End Function

''
' Create Sheet on FileName
'
' @method createNewSheet
' @param {String} 시트 명
' @param {XlThemeColor|Empty} xlTheme 테마 컬러
' @return {WorkSheet} 추가된 시트
''
Public Function createSheet(ByVal sheetName As String, Optional xlTheme As Excel.XlThemeColor = 7) As Worksheet
    Dim sheet As Worksheet
        
    If Not isSheet(sheetName) Then
        ' 시트가 존재하지 않는경우 생성
        Set sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        
        ' 시트이름 설정 및, 시트 탭 색상 설정
        With sheet
            .Name = sheetName
            With .Tab
                .ThemeColor = xlTheme
                .TintAndShade = 0.399975585192419
            End With
            .Activate
        End With
    Else
        ' 시트가 존재하는경우 취득
        Debug.Print "Already Exists Sheet " + sheetName
        Set sheet = ActiveWorkbook.Worksheets(sheetName)
        sheet.Activate
    End If
        
    ' Return
    Set createSheet = sheet
End Function
