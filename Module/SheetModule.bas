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
' @param {String} ��Ʈ��
' @return {Boolean} ��Ʈ ���� ����
''
Public Function isSheet(ByVal sheetName As String) As Boolean
    ' ���� �߻��� ������ ��ġ ����
    On Error GoTo e0
    ' ��Ʈ�� ��� ��, ��Ʈ���� �����ϴ� ���, True
    If CBool(Len(Sheets(sheetName).Name)) Then isSheet = True
    Exit Function
    
e0:
    isSheet = False
End Function

''
' Remove Sheet For Name
'
' @method removeSheet
' @param {String} ��Ʈ��
' @return {Integer} -1: ������ / 0: �������� / 1: �̻�����
''
Public Function removeSheet(ByVal sheetName As String) As Integer
    Dim wbs As Worksheet
    ' ���̾�α� ��ǥ��
    Application.DisplayAlerts = False
    
    ' ���� �߻��� ���� ��ġ
    On Error GoTo e0
    
    ' ��Ʈ�� �˻�
    For Each wbs In Sheets
        With wbs
            If .Name = sheetName Then
                .Delete
                ' ������ ���
                removeSheet = 0
                GoTo s0
            End If
        End With
    Next
    
    ' ������ ����� ���°��
    removeSheet = -1
s0:
    ' ���� ����
    Application.DisplayAlerts = True
    Exit Function
e0:
    ' �̻� ����
    removeSheet = 1
End Function

''
' Create Sheet on FileName
'
' @method createNewSheet
' @param {String} ��Ʈ ��
' @param {XlThemeColor|Empty} xlTheme �׸� �÷�
' @return {WorkSheet} �߰��� ��Ʈ
''
Public Function createSheet(ByVal sheetName As String, Optional xlTheme As Excel.XlThemeColor = 7) As Worksheet
    Dim sheet As Worksheet
        
    If Not isSheet(sheetName) Then
        ' ��Ʈ�� �������� �ʴ°�� ����
        Set sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        
        ' ��Ʈ�̸� ���� ��, ��Ʈ �� ���� ����
        With sheet
            .Name = sheetName
            With .Tab
                .ThemeColor = xlTheme
                .TintAndShade = 0.399975585192419
            End With
            .Activate
        End With
    Else
        ' ��Ʈ�� �����ϴ°�� ���
        Debug.Print "Already Exists Sheet " + sheetName
        Set sheet = ActiveWorkbook.Worksheets(sheetName)
        sheet.Activate
    End If
        
    ' Return
    Set createSheet = sheet
End Function
