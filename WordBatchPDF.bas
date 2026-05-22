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
    Dim fd As FileDialog
    Dim folderPath As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim doc As Document
    Dim ext As String
    Dim baseName As String
    Dim pdfPath As String
    Dim success As Long
    Dim total As Long
    Dim failed As String
    Dim msg As String
    
    ' 1. Select folder
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Select folder with Word files to convert to PDF"
    If fd.Show <> -1 Then Exit Sub
    folderPath = fd.SelectedItems(1)
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' 2. Validate folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Folder does not exist." & vbCrLf & folderPath, vbExclamation, "Word Batch PDF"
        Exit Sub
    End If
    
    Set folder = fso.GetFolder(folderPath)
    success = 0
    total = 0
    failed = ""
    
    Application.ScreenUpdating = False
    
    ' 3. Process each file
    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.Path))
        If IsWordExtension(ext) Then
            total = total + 1
            Set doc = Nothing
            On Error Resume Next
            
            ' Open document read-only and hidden
            ' Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible)
            Set doc = Documents.Open(file.Path, , True, , , , , , , , , False)
            If doc Is Nothing Then
                failed = failed & file.Name & " - " & Err.Description & vbCrLf
                Err.Clear
            Else
                ' Build PDF path (same folder, .pdf extension)
                baseName = fso.GetBaseName(file.Path)
                pdfPath = folderPath & baseName & ".pdf"
                
                ' Export to PDF — positional arguments for maximum compatibility
                ' ExportAsFixedFormat(OutputFileName, ExportFormat, [OpenAfterExport], [OptimizeFor], [Range])
                doc.ExportAsFixedFormat pdfPath, wdExportFormatPDF, False, wdExportOptimizeForPrint, wdExportAllDocument
                
                If Err.Number <> 0 Then
                    failed = failed & file.Name & " - " & Err.Description & vbCrLf
                    Err.Clear
                Else
                    success = success + 1
                End If
                
                ' Close without saving changes to original
                doc.Close wdDoNotSaveChanges
            End If
            On Error GoTo 0
        End If
    Next file
    
    Application.ScreenUpdating = True
    
    ' 4. Show summary
    If total = 0 Then
        MsgBox "No Word files found in the selected folder." & vbCrLf & _
               "Supported: .doc .docx .docm .dot .dotx .dotm .odt .rtf", _
               vbInformation, "Word Batch PDF"
    ElseIf failed = "" Then
        msg = "Conversion complete!" & vbCrLf & _
              "All " & success & " files converted to PDF successfully." & vbCrLf & vbCrLf & _
              "Output folder: " & folderPath
        MsgBox msg, vbInformation, "Word Batch PDF"
    Else
        msg = "Conversion finished with errors." & vbCrLf & vbCrLf & _
              "Successful: " & success & " / " & total & vbCrLf & vbCrLf & _
              "Failed files:" & vbCrLf & failed
        MsgBox msg, vbExclamation, "Word Batch PDF"
    End If
    
    ' 5. Open output folder in Explorer
    On Error Resume Next
    Shell "explorer.exe """ & folderPath & """", vbNormalFocus
End Sub

'=====================================================================
' Check if extension is a Word-compatible format
'=====================================================================
Private Function IsWordExtension(ByVal ext As String) As Boolean
    Select Case ext
        Case "doc", "docx", "docm", "dot", "dotx", "dotm", "odt", "rtf"
            IsWordExtension = True
        Case Else
            IsWordExtension = False
    End Select
End Function

'=====================================================================
' End of Module
'=====================================================================
