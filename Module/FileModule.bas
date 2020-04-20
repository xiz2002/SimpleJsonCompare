Attribute VB_Name = "FileModule"
''
' Json Simple Compare
' https://github.com/xiz2002/SimpleJsonCompare
'
' FileModule
'
' @class FileModule
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
' Dialog Open get File Path
'
' @method OpenDialog
' @return {String} SelectedFilePath
''
Public Function openDialog() As String
    ' Return FilePath Value
    Dim filePath As String
    
    ' File Dialog Variable
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ' File Dialog
    With fd
        .Filters.Clear
        .Title = ""
        .Filters.Add "", "*.json", 1
        .AllowMultiSelect = False
        
        ' Pressed Open Button
        If .Show = -1 Then filePath = .SelectedItems(1)
    End With
     
     'Set the object variable to nothing.
    Set fd = Nothing
    
    ' Return FilePath
    openDialog = filePath
End Function

''
' Read File For filePath
'
' @method readText
' @param {String} filePath
' @return {String} JsonText
''
Public Function readText(ByVal filePath) As String
    ' File System, File, FilePath
    Dim fso As Scripting.FileSystemObject, file As file
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' TextStream, String
    Dim ts As TextStream, jsonText As String
    
    ' Check Is Null the File Path
    If Trim(filePath & vbNullString) = vbNullString Then
        Exit Function
    End If
    
    ' GetFile
    Set file = fso.GetFile(filePath)

    ' Open Stream
    Set ts = file.OpenAsTextStream(ForReading, TristateUseDefault)
    
    ' Read File
    jsonText = ts.ReadAll

    ' Close Stream
    ts.Close
    
    'Set the object variable to nothing.
    Set file = Nothing
    Set fso = Nothing
    
    'Return FileText
    readText = jsonText
End Function
