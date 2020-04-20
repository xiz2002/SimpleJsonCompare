VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SimpleCompForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13365
   OleObjectBlob   =   "SimpleCompForm.frx":0000
   StartUpPosition =   1  '������ ���
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

' ��������
Private m_JsonTableList As Collection                                                   ' JsonToTable Object Collection

'' Constructor''
Private Sub UserForm_Initialize()
    ' �ݷ��� ����
    Set m_JsonTableList = New Collection
    ' ������ ������� ����
    Application.Calculation = xlManual
End Sub

'' Destory ''
Private Sub UserForm_Terminate()
    On Error Resume Next
    ' �ݷ��� ���� ��ü ����
    Call collectionCleaner
    ' �ݷ��� ����
    Set m_JsonTableList = Nothing
    ' ������ �ڵ���� ����
    Application.Calculation = xlAutomatic
End Sub

''
' Collection Cleaner
' �ݷ��� ���� ������Ʈ ����ó��
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
' �����ư �̺�Ʈ
'
' @method btnExecute_Click
''
Private Sub btnExecute_Click()
    Dim ctrl As Control, instance As JsonToTable
    Dim ctrlTag As String, jsonData As String
    Dim tableNames As Variant, index As Long
    
    ' Clean for Collection
    Call collectionCleaner
    
    ' �� ��Ʈ�� �˻�
    For Each ctrl In Me.Controls
        ' �ؽ�Ʈ �ڽ��� ���, �ؽ�Ʈ ���� ���� ��, ��ü ����
        If TypeOf ctrl Is MSForms.TextBox Then
            ' �±� ���
            ctrlTag = ctrl.Tag
        
            If Trim(ctrl.Text & vbNullString) = vbNullString And Not ctrl.Name = "tbFilePath3" Then
                ' File_3 �̿� ���� ��ΰ� ���� ���� �ʴ°��, e0���� ����
                GoTo e0
            ElseIf Trim(ctrl.Text & vbNullString) = vbNullString And ctrl.Name = "tbFilePath3" Then
                ' File_3 ���� ��ΰ� ���� ���� �ʴ°��
                ' ���� ����
                Call QueryModule.removeQuery(ctrlTag)
                ' Transpose���� ����
                Call QueryModule.removeQuery(QueryModule.PREFIX_TRANSPOSE + ctrlTag)
                ' ��Ʈ ����
                Call SheetModule.removeSheet(ctrlTag)
                ' �ݺ��� ����
                Exit For
            End If
            
            ' ���ϰ�θ� �̿���, Json ���� ���� ���
            jsonData = FileModule.readText(ctrl.Text)
            ' fileName = Mid(Dir(ctrl.Text), 1, InStrRev(LCase(Dir(ctrl.Text)), ".") - 1)
            
            ' Json���� ������ ���� ���� �ʴ°��, e1���� ����
            If Trim(jsonData & vbNullString) = vbNullString Then
                GoTo e1
            End If
            
            ' JsonToTable Ŭ���� ���� �� �ʱ�ȭ
            Set instance = New JsonToTable
            instance.jsonData = jsonData
            instance.tableName = ctrlTag
            ' �ν��Ͻ��� �ݷ��ǿ� �߰�
            m_JsonTableList.Add item:=instance, key:=ctrlTag
            ' �޸� ����
            Set instance = Nothing
            
            If IsEmpty(tableNames) Then
                ' ���̺� �̸� �迭 ����
                tableNames = Array(ctrlTag)
            ElseIf UBound(tableNames) >= 0 Then
                ' ���̺� �̸� �迭�� �����ϸ�, �迭 ������ ���� ��, ���̺� �̸� �߰�
                index = UBound(tableNames) + 1
                ReDim Preserve tableNames(index)
                tableNames(index) = ctrlTag
            End If

        End If
    Next
    
    ' Json ���� ������ ���̺�� ����
    For Each instance In m_JsonTableList
        Call instance.ConvertJsonToTable
    Next
    
    ' Query ����
    For Each instance In m_JsonTableList
        Call QueryModule.createQuery(instance.tableName)
    Next
    
    ' Qeury Combine
    Call QueryModule.createCombineQuery("Compare", tableNames)
    
    ' ������ �ڵ���� ����
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
