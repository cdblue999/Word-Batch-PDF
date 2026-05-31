' Copyright (C) 2026 ZMS
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <https://www.gnu.org/licenses/>.
Attribute VB_Name = "WordBatchPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=====================================================================
' Word Batch PDF Converter
' Microsoft Word 2013 VBA macro
'
' Converts all Word files in a selected folder to PDF.
'
' Supported formats:
'   - .doc, .docx, .docm (Word documents)
'   - .dot, .dotx, .dotm (Word templates)
'   - .odt, .rtf (Word-compatible formats)
'
' HOW TO INSTALL:
'
' 1. Open Word 2013
' 2. Press Alt+F11 to open the VBA editor
' 3. File > Import File > select WordBatchPDF.bas
' 4. Close the VBA editor
'
' HOW TO USE:
'
' 1. View > Macros > View Macros (or Alt+F8)
' 2. Select WordBatchPDF.BatchWordToPDF
' 3. Click Run
' 4. Select a folder with Word files
' 5. Macro processes all files and saves PDFs in the same folder
' 6. Output folder opens automatically in Explorer
'=====================================================================

'=====================================================================
' Main entry point
'=====================================================================
Public Sub BatchWordToPDF()
    Dim fd As Object
    Dim folderPath As String
    Dim fso As Object
    Dim folder As Object
    Dim f As Object
    Dim doc As Object
    Dim ext As String
    Dim baseName As String
    Dim pdfPath As String
    Dim okCount As Long
    Dim errCount As Long
    Dim errList As String
    Dim msg As String
    
    On Error GoTo ErrHandler
    
    ' 1. Select folder
    Set fd = Application.FileDialog(3)
    fd.Title = "Select folder with Word files to convert to PDF"
    If fd.Show <> -1 Then Exit Sub
    folderPath = fd.SelectedItems(1)
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' 2. Validate folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) = False Then
        MsgBox "Folder does not exist.", vbExclamation, "Word Batch PDF"
        Exit Sub
    End If
    
    Set folder = fso.GetFolder(folderPath)
    okCount = 0
    errCount = 0
    errList = ""
    
    Application.ScreenUpdating = False
    
    ' 3. Process each file
    For Each f In folder.Files
        ext = LCase(fso.GetExtensionName(f.Path))
        If IsWordExt(ext) Then
            On Error Resume Next
            Set doc = Nothing
            Set doc = Documents.Open(f.Path, , True)
            If doc Is Nothing Then
                errCount = errCount + 1
                errList = errList & f.Name & vbCrLf
                Err.Clear
            Else
                baseName = fso.GetBaseName(f.Path)
                pdfPath = folderPath & baseName & ".pdf"
                doc.ExportAsFixedFormat pdfPath, 0, False, 0, 0
                If Err.Number <> 0 Then
                    errCount = errCount + 1
                    errList = errList & f.Name & vbCrLf
                    Err.Clear
                Else
                    okCount = okCount + 1
                End If
                doc.Close 0
            End If
            On Error GoTo 0
        End If
    Next f
    
    Application.ScreenUpdating = True
    
    ' 4. Show summary
    If okCount = 0 And errCount = 0 Then
        MsgBox "No Word files found in the selected folder.", vbInformation, "Word Batch PDF"
    ElseIf errCount = 0 Then
        msg = "Conversion complete!" & vbCrLf & vbCrLf & _
              "All " & okCount & " files converted to PDF.", vbInformation, "Word Batch PDF"
        MsgBox msg, vbInformation, "Word Batch PDF"
    ElseIf okCount > 0 Then
        msg = "Partial success:" & vbCrLf & vbCrLf & _
              "OK: " & okCount & "  Failed: " & errCount & vbCrLf & vbCrLf & _
              "Errors:" & vbCrLf & errList
        MsgBox msg, vbExclamation, "Word Batch PDF"
    Else
        msg = "All files failed:" & vbCrLf & vbCrLf & errList
        MsgBox msg, vbExclamation, "Word Batch PDF"
    End If
    
    Shell "explorer.exe """ & folderPath & """", 1
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical, "Word Batch PDF"
End Sub

'=====================================================================
' Check if extension is a Word-compatible format
'=====================================================================
Private Function IsWordExt(ByVal ext As String) As Boolean
    Select Case ext
        Case "doc", "docx", "docm", "dot", "dotx", "dotm", "odt", "rtf"
            IsWordExt = True
        Case Else
            IsWordExt = False
    End Select
End Function
