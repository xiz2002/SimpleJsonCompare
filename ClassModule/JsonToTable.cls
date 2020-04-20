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

' ��������
Private m_JsonData As String                                    ' Json ���� ����
Private m_TableName As String                                   ' Table ��
Private m_Table As ListObject                                   ' Table ��ü
Private m_Data As Object                                        ' Json ���� ���� �ؼ� ��� (Dictionary or Collection)

''
' JsonData(RAW) Setter
' @param {String} data JsonRawData
''
Property Let jsonData(data As String)
    m_JsonData = data
End Property

''
' Getter TableName
' @param {String} tableName ���̺� ��
''
Property Get tableName() As String
    tableName = m_TableName
End Property

''
' Setter TableName
' @param {String} tableName ���̺� ��
''
Property Let tableName(tableName As String)
    m_TableName = tableName
End Property

'' ������''
Private Sub Class_Initialize()
End Sub

'' �ı���''
Private Sub Class_Terminate()
    Set m_Table = Nothing
    Set m_Data = Nothing
End Sub

''
' Json Data To Table
'
' @method ConvertJsonToTable
' @param {String|Empty} Json���� ����
' @param {String|Empty} Table �̸�
''
Public Sub ConvertJsonToTable(Optional ByVal jsonData = "", Optional ByVal tableName = "")
    ' �μ��� �����ϴ� ���, Json ���� ���� ����
    If Not Trim(jsonData & vbNullString) = vbNullString Then
        m_JsonData = jsonData
    End If
    
    ' �μ��� �����ϴ� ���, ���̺� �̸� ����
    If Not Trim(tableName & vbNullString) = vbNullString Then
        m_TableName = tableName
    End If
    
    ' ���� ���� �Ǵ� ���̺� �̸��� �������� �ʴ°�� ó�� ����
     If Trim(m_JsonData & vbNullString) = vbNullString Or Trim(m_TableName & vbNullString) = vbNullString Then
        Exit Sub
    End If
        
    ' ���̺� �ۼ�
    Set m_Table = TableModule.createTable(SheetModule.createSheet(m_TableName), m_TableName)
    
    ' Json ���� ���� ���
    Set m_Data = JsonConverter.ParseJson(m_JsonData)
    
    ' Json ���� ���� ����
    Call parsedDataDivision(m_Data)
End Sub

''
' Json Data Type Division
'
' @method parsedDataDivision
' @param {Dictionary|Collection} obj JsonParsing������
' @param {String|Empty} pField �����ʵ��
''
Private Sub parsedDataDivision(ByVal obj As Object, Optional ByVal pField As String = "")
    If VBA.TypeName(obj) = "Dictionary" Then
        ' �ڷ�Ÿ���� �����ΰ��
        Dim key As Variant
        
        ' Key ���� ��ŭ ����
        For Each key In obj.Keys
            Call divDataAndInsert(obj(key), _
                                IIf(Trim(pField & vbNullString) = vbNullString, key, pField & "." & key))
        Next
    ElseIf VBA.TypeName(obj) = "Collection" Then
        ' �ڷ�Ÿ���� �ݷ����ΰ��
        Dim index As Integer
        
        ' �ε��� �� ��ŭ ����
        For index = 1 To obj.Count
            Call divDataAndInsert(obj(index), _
                                IIf(Trim(pField & vbNullString) = vbNullString, CStr(index), pField & "." & CStr(index)))
        Next index
    Else
        ' �׿� �ڷ�Ÿ��
        Debug.Print "Type Miss:" + pField + CStr(obj)
    End If
End Sub

''
' Json Item Division And Write Table
'
' @method divDataAndInsert
' @param {Variant} data ��
' @param {String} field �ʵ�
''
Private Function divDataAndInsert(ByVal data As Variant, ByVal field As String)
    Select Case VBA.VarType(data)
    Case VBA.vbNull
        ' �ڷ�Ÿ���� Null�ΰ��
        ' Debug.Print "[Field]:" + CStr(field) + " [Value]:" + "Null" + " [Type]: Null"
        Call TableModule.addDataToTable(m_Table, Array(CStr(field), Null))
    Case VBA.vbDate
        ' �ڷ�Ÿ���� Date�ΰ��
        ' Debug.Print "[Field]:" + CStr(field) + " [Value]:" + CStr(data) + " [Type]: Date"
        Call TableModule.addDataToTable(m_Table, Array(CStr(field), data))
    Case VBA.vbString
        ' �ڷ�Ÿ���� ���ڿ��ΰ��
        ' Debug.Print "[Field]:" + CStr(field) + " [Value]:" + CStr(data) + " [Type]: String"
        Call TableModule.addDataToTable(m_Table, Array(CStr(field), CStr("'" + data)))
    Case VBA.vbBoolean
        ' �ڷ�Ÿ���� Bool�ΰ��
        ' Debug.Print "[Field]:" + CStr(field) + " [Value]:" + CStr(data) + " [Type]: Boolean"
        Call TableModule.addDataToTable(m_Table, Array(CStr(field), data))
    Case VBA.vbArray
        ' �ڷ�Ÿ���� �迭�ΰ��
        ' Debug.Print "[Field]:" + CStr(field) + " [Value]:" + CStr(data) + " [Type]: Array"
        Call TableModule.addDataToTable(m_Table, Array(CStr(field), data))
    Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
        ' �ڷ�Ÿ���� �����ΰ��
        ' Debug.Print "[Field]:" + CStr(field) + " [Value]:" + CStr(data) + " [Type]: Number"
        Call TableModule.addDataToTable(m_Table, Array(CStr(field), data))
    Case VBA.vbObject
        If VBA.TypeName(data) = "Dictionary" Then
            ' �ڷ�Ÿ���� �����ΰ��
            
            ' Empty Check
            If Not CBool(data.Count) Then
                ' Debug.Print "[Field]:" + CStr(field) + " [Value]:" + "Empty" + " [Type]: Dictionary"
                Call TableModule.addDataToTable(m_Table, Array(CStr(field), Empty))
                Exit Function
            End If
            
            Call parsedDataDivision(data, field)
        ElseIf VBA.TypeName(data) = "Collection" Then
            ' �ڷ�Ÿ���� �ݷ����ΰ��
            
            ' Empty Check
            If Not CBool(data.Count) Then
                ' Debug.Print "[Field]:" + CStr(field) + " [Value]:" + "Empty" + " [Type]: Collection"
                Call TableModule.addDataToTable(m_Table, Array(CStr(field), Empty))
                Exit Function
            End If
            
            ' Note: ���� ����Ʈ �� ���� �� �ʿ䰡 �ִ°�?
            If VBA.VarType(data(1)) = VBA.vbObject Then
                ' �ݷ��� ���� �ڷ����� ������Ʈ�� ���
                Call parsedDataDivision(data, field)
            Else
                ' �ݷ��� ���� �ڷ����� ������Ʈ �̿��� ���
                Dim accItem As String, item As Variant
                ' ���� �ڷ� ��� �� ���ڿ��� ����
                For Each item In data
                    accItem = IIf(Trim(accItem & vbNullString) = vbNullString, CStr(item), accItem & "," & CStr(item))
                Next
                
                Call TableModule.addDataToTable(m_Table, Array(CStr(field), accItem))
            End If
        End If
    End Select
End Function