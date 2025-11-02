Attribute VB_Name = "Module1_Core"
'=============================================================================
' Module1_Core.bas
' Main export orchestration for QRVCARD Generator V4
'
' Purpose: Coordinate export operations and user interface
'
' Public Functions:
'   - ExportSingleRow()       Export current row (vCard + HTML)
'   - ExportVcardOnly()       Export current row vCard only
'   - ExportHTMLOnly()        Export current row HTML only
'   - ExportAll()             Batch export marked rows WITH PROGRESS BAR
'
' Dependencies:
'   - Module2_VCardBuilder (vCard generation)
'   - Module3_HTMLBuilder (HTML generation)
'   - Module4_QRCode (QR code generation)
'   - Module5_Utilities (helper functions)
'   - frmProgress (MANDATORY progress bar for batch export)
'
' CRITICAL: Progress bar is NON-OPTIONAL for batch exports
'=============================================================================

Option Explicit

'-----------------------------------------------------------------------------
' ExportSingleRow
' Purpose: Export current row as both vCard and HTML
' User Interface: Attached to "Export Single Row" button
'-----------------------------------------------------------------------------
Public Sub ExportSingleRow()
    Dim ws As Worksheet
    Dim wsSettings As Worksheet
    Dim currentRow As Long
    Dim folderPath As String
    Dim fileName As String
    Dim lastName As String, firstName As String
    Dim vcfPath As String, htmlPath As String
    Dim vcfSuccess As Boolean, htmlSuccess As Boolean
    
    Set ws = ThisWorkbook.Sheets("Data Entry")
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    
    ' Get current row
    currentRow = ActiveCell.Row
    
    If currentRow < 2 Then
        MsgBox "Please select a data row (row 2 or below).", vbExclamation, "Invalid Selection"
        Exit Sub
    End If
    
    ' Validate required fields
    If Not ValidateRequiredFields(ws, currentRow) Then
        Exit Sub
    End If
    
    ' Get folder path
    folderPath = GetFolderPath()
    If folderPath = "" Then Exit Sub
    
    ' Generate base filename
    lastName = Module5_Utilities.GetCellValue(ws, currentRow, "B")
    firstName = Module5_Utilities.GetCellValue(ws, currentRow, "C")
    fileName = Module5_Utilities.SanitizeFilename(lastName & "_" & firstName)
    
    vcfPath = folderPath & fileName & ".vcf"
    htmlPath = folderPath & fileName & ".html"
    
    ' Export vCard
    vcfSuccess = Module2_VCardBuilder.SaveVcard(ws, currentRow, vcfPath)
    
    ' Export HTML
    htmlSuccess = Module3_HTMLBuilder.SaveHTML(ws, wsSettings, currentRow, htmlPath)
    
    ' Update tracking
    If vcfSuccess And htmlSuccess Then
        LogExport ws, currentRow, "BOTH"
        MsgBox "Export successful!" & vbCrLf & vbCrLf & _
               "vCard: " & vcfPath & vbCrLf & _
               "HTML: " & htmlPath, _
               vbInformation, "Success"
    Else
        MsgBox "Export failed. Check file permissions and try again.", vbExclamation, "Error"
    End If
End Sub

'-----------------------------------------------------------------------------
' ExportVcardOnly
' Purpose: Export current row as vCard only
' User Interface: Attached to "Export vCard Only" button
'-----------------------------------------------------------------------------
Public Sub ExportVcardOnly()
    Dim ws As Worksheet
    Dim currentRow As Long
    Dim folderPath As String
    Dim fileName As String
    Dim lastName As String, firstName As String
    Dim vcfPath As String
    
    Set ws = ThisWorkbook.Sheets("Data Entry")
    
    ' Get current row
    currentRow = ActiveCell.Row
    
    If currentRow < 2 Then
        MsgBox "Please select a data row (row 2 or below).", vbExclamation, "Invalid Selection"
        Exit Sub
    End If
    
    ' Validate required fields
    If Not ValidateRequiredFields(ws, currentRow) Then
        Exit Sub
    End If
    
    ' Get folder path
    folderPath = GetFolderPath()
    If folderPath = "" Then Exit Sub
    
    ' Generate filename
    lastName = Module5_Utilities.GetCellValue(ws, currentRow, "B")
    firstName = Module5_Utilities.GetCellValue(ws, currentRow, "C")
    fileName = Module5_Utilities.SanitizeFilename(lastName & "_" & firstName)
    vcfPath = folderPath & fileName & ".vcf"
    
    ' Export vCard
    If Module2_VCardBuilder.SaveVcard(ws, currentRow, vcfPath) Then
        LogExport ws, currentRow, "VCARD"
        MsgBox "vCard exported successfully!" & vbCrLf & vbCrLf & vcfPath, vbInformation, "Success"
    Else
        MsgBox "vCard export failed. Check file permissions and try again.", vbExclamation, "Error"
    End If
End Sub

'-----------------------------------------------------------------------------
' ExportHTMLOnly
' Purpose: Export current row as HTML only
' User Interface: Attached to "Export HTML Only" button
'-----------------------------------------------------------------------------
Public Sub ExportHTMLOnly()
    Dim ws As Worksheet
    Dim wsSettings As Worksheet
    Dim currentRow As Long
    Dim folderPath As String
    Dim fileName As String
    Dim lastName As String, firstName As String
    Dim htmlPath As String
    
    Set ws = ThisWorkbook.Sheets("Data Entry")
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    
    ' Get current row
    currentRow = ActiveCell.Row
    
    If currentRow < 2 Then
        MsgBox "Please select a data row (row 2 or below).", vbExclamation, "Invalid Selection"
        Exit Sub
    End If
    
    ' Validate required fields
    If Not ValidateRequiredFields(ws, currentRow) Then
        Exit Sub
    End If
    
    ' Get folder path
    folderPath = GetFolderPath()
    If folderPath = "" Then Exit Sub
    
    ' Generate filename
    lastName = Module5_Utilities.GetCellValue(ws, currentRow, "B")
    firstName = Module5_Utilities.GetCellValue(ws, currentRow, "C")
    fileName = Module5_Utilities.SanitizeFilename(lastName & "_" & firstName)
    htmlPath = folderPath & fileName & ".html"
    
    ' Export HTML
    If Module3_HTMLBuilder.SaveHTML(ws, wsSettings, currentRow, htmlPath) Then
        LogExport ws, currentRow, "HTML"
        MsgBox "HTML exported successfully!" & vbCrLf & vbCrLf & htmlPath, vbInformation, "Success"
    Else
        MsgBox "HTML export failed. Check file permissions and try again.", vbExclamation, "Error"
    End If
End Sub

'-----------------------------------------------------------------------------
' ExportAll
' Purpose: Batch export all marked rows (Column A has value)
' User Interface: Attached to "Export All Marked Rows" button
' CRITICAL: Uses MANDATORY progress bar (frmProgress)
'-----------------------------------------------------------------------------
Public Sub ExportAll()
    Dim ws As Worksheet
    Dim wsSettings As Worksheet
    Dim markedRows As Collection
    Dim rowNum As Long
    Dim folderPath As String
    Dim lastName As String, firstName As String
    Dim fileName As String
    Dim vcfPath As String, htmlPath As String
    Dim successCount As Long, failCount As Long
    Dim i As Long
    Dim progressForm As frmProgress
    
    Set ws = ThisWorkbook.Sheets("Data Entry")
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    
    ' Get marked rows (Column A has value)
    Set markedRows = GetMarkedRows(ws)
    
    If markedRows.Count = 0 Then
        MsgBox "No rows selected for export." & vbCrLf & vbCrLf & _
               "Mark rows in Column A (type 'x' or any character) to select them for export.", _
               vbExclamation, "No Selection"
        Exit Sub
    End If
    
    ' Get folder path
    folderPath = GetFolderPath()
    If folderPath = "" Then Exit Sub
    
    ' === MANDATORY PROGRESS BAR ===
    Set progressForm = New frmProgress
    progressForm.Show vbModeless
    progressForm.InitializeProgress markedRows.Count
    
    successCount = 0
    failCount = 0
    
    ' Process each marked row
    For i = 1 To markedRows.Count
        rowNum = markedRows(i)
        
        ' Update progress bar
        firstName = Module5_Utilities.GetCellValue(ws, rowNum, "C")
        lastName = Module5_Utilities.GetCellValue(ws, rowNum, "B")
        progressForm.UpdateProgress i, lastName & ", " & firstName
        DoEvents
        
        ' Validate required fields
        If Not ValidateRequiredFields(ws, rowNum, False) Then
            failCount = failCount + 1
            GoTo NextRow
        End If
        
        ' Generate base filename
        fileName = Module5_Utilities.SanitizeFilename(lastName & "_" & firstName)
        vcfPath = folderPath & fileName & ".vcf"
        htmlPath = folderPath & fileName & ".html"
        
        ' Export vCard and HTML
        If Module2_VCardBuilder.SaveVcard(ws, rowNum, vcfPath) Then
            If Module3_HTMLBuilder.SaveHTML(ws, wsSettings, rowNum, htmlPath) Then
                LogExport ws, rowNum, "BOTH"
                successCount = successCount + 1
            Else
                failCount = failCount + 1
            End If
        Else
            failCount = failCount + 1
        End If
        
NextRow:
    Next i
    
    ' Close progress bar
    Unload progressForm
    
    ' Show summary
    MsgBox "Batch export complete!" & vbCrLf & vbCrLf & _
           "Successful: " & successCount & vbCrLf & _
           "Failed: " & failCount & vbCrLf & vbCrLf & _
           "Files saved to: " & folderPath, _
           vbInformation, "Export Complete"
End Sub

'-----------------------------------------------------------------------------
' GetMarkedRows
' Purpose: Get collection of row numbers marked for export (Column A has value)
' Parameters:
'   ws - Data Entry worksheet
' Returns: Collection of row numbers
'-----------------------------------------------------------------------------
Private Function GetMarkedRows(ByVal ws As Worksheet) As Collection
    Dim markedRows As New Collection
    Dim lastRow As Long
    Dim i As Long
    Dim markerValue As String
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow  ' Start at row 2 (skip header)
        markerValue = Trim(ws.Cells(i, "A").value)
        If markerValue <> "" Then
            markedRows.Add i
        End If
    Next i
    
    Set GetMarkedRows = markedRows
End Function

'-----------------------------------------------------------------------------
' ValidateRequiredFields
' Purpose: Check that all required fields are filled
' Parameters:
'   ws - Data Entry worksheet
'   rowNum - Row number to validate
'   showMessage - Show error message if True (default True)
' Returns: True if valid, False if missing required fields
'-----------------------------------------------------------------------------
Private Function ValidateRequiredFields(ByVal ws As Worksheet, ByVal rowNum As Long, Optional ByVal showMessage As Boolean = True) As Boolean
    Dim missingFields As String
    Dim value As String
    
    missingFields = ""
    
    ' Required fields (marked with * in header)
    ' B: Last Name
    value = Module5_Utilities.GetCellValue(ws, rowNum, "B")
    If value = "" Then missingFields = missingFields & "Last Name, "
    
    ' C: First Name
    value = Module5_Utilities.GetCellValue(ws, rowNum, "C")
    If value = "" Then missingFields = missingFields & "First Name, "
    
    ' F: Email
    value = Module5_Utilities.GetCellValue(ws, rowNum, "F")
    If value = "" Then missingFields = missingFields & "Email, "
    
    ' G: Phone Cell
    value = Module5_Utilities.GetCellValue(ws, rowNum, "G")
    If value = "" Then missingFields = missingFields & "Phone Cell, "
    
    ' J: Title Primary
    value = Module5_Utilities.GetCellValue(ws, rowNum, "J")
    If value = "" Then missingFields = missingFields & "Title Primary, "
    
    ' L: Organization Primary
    value = Module5_Utilities.GetCellValue(ws, rowNum, "L")
    If value = "" Then missingFields = missingFields & "Organization Primary, "
    
    ' O: City
    value = Module5_Utilities.GetCellValue(ws, rowNum, "O")
    If value = "" Then missingFields = missingFields & "City, "
    
    ' P: State
    value = Module5_Utilities.GetCellValue(ws, rowNum, "P")
    If value = "" Then missingFields = missingFields & "State, "
    
    ' Q: Postal Code
    value = Module5_Utilities.GetCellValue(ws, rowNum, "Q")
    If value = "" Then missingFields = missingFields & "Postal Code, "
    
    ' Check if any fields missing
    If missingFields <> "" Then
        ' Remove trailing comma and space
        missingFields = Left(missingFields, Len(missingFields) - 2)
        
        If showMessage Then
            MsgBox "Missing required fields in row " & rowNum & ":" & vbCrLf & vbCrLf & _
                   missingFields & vbCrLf & vbCrLf & _
                   "Please fill in all required fields (marked with * in header).", _
                   vbExclamation, "Validation Error"
        End If
        
        ValidateRequiredFields = False
    Else
        ValidateRequiredFields = True
    End If
End Function

'-----------------------------------------------------------------------------
' GetFolderPath
' Purpose: Show folder picker dialog
' Returns: Selected folder path with trailing slash, empty string if cancelled
'-----------------------------------------------------------------------------
Private Function GetFolderPath() As String
    Dim folderPicker As FileDialog
    Dim selectedPath As String
    
    Set folderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    
    With folderPicker
        .Title = "Select Export Folder"
        .AllowMultiSelect = False
        .InitialFileName = Environ("USERPROFILE") & "\Documents\"
        
        If .Show = -1 Then
            selectedPath = .SelectedItems(1)
            
            ' Ensure trailing slash
            If Right(selectedPath, 1) <> "\" Then
                selectedPath = selectedPath & "\"
            End If
            
            GetFolderPath = selectedPath
        Else
            GetFolderPath = ""
        End If
    End With
End Function

'-----------------------------------------------------------------------------
' LogExport
' Purpose: Update export tracking columns (AI-AL)
' Parameters:
'   ws - Data Entry worksheet
'   rowNum - Row number
'   exportType - "VCARD", "HTML", or "BOTH"
'-----------------------------------------------------------------------------
Private Sub LogExport(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal exportType As String)
    On Error Resume Next
    
    Dim currentCount As Long
    
    ' AI: vCard Exported
    If exportType = "VCARD" Or exportType = "BOTH" Then
        ws.Cells(rowNum, "AI").value = "TRUE"
    End If
    
    ' AJ: HTML Exported
    If exportType = "HTML" Or exportType = "BOTH" Then
        ws.Cells(rowNum, "AJ").value = "TRUE"
    End If
    
    ' AK: Last Export Date
    ws.Cells(rowNum, "AK").value = Now
    ws.Cells(rowNum, "AK").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    
    ' AL: Export Count (increment)
    currentCount = 0
    If IsNumeric(ws.Cells(rowNum, "AL").value) Then
        currentCount = CLng(ws.Cells(rowNum, "AL").value)
    End If
    ws.Cells(rowNum, "AL").value = currentCount + 1
    
    On Error GoTo 0
End Sub
