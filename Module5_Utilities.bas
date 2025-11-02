Attribute VB_Name = "Module5_Utilities"
'=============================================================================
' Module5_Utilities.bas
' Shared utility functions for QRVCARD Generator V4
'
' Purpose: Provide reusable helper functions to eliminate code duplication
'
' Public Functions:
'   - SanitizeFilename()      Clean filenames for safe file system use
'   - GetCellValue()          Safe cell value retrieval with error handling
'   - BuildFormattedName()    Construct full name from components
'   - GetSettingValue()       Retrieve setting with fallback default
'   - DebugLog()              Centralized debug logging
'
' Dependencies: None (standalone utility module)
'=============================================================================

Option Explicit

'-----------------------------------------------------------------------------
' SanitizeFilename
' Purpose: Remove invalid filesystem characters from filename
' Returns: Safe filename string
'-----------------------------------------------------------------------------
Public Function SanitizeFilename(ByVal filename As String) As String
    Dim invalidChars As String
    Dim i As Integer
    Dim result As String
    
    ' Invalid characters for Windows filenames
    invalidChars = "\/:*?""<>|"
    
    result = filename
    
    ' Replace each invalid character with underscore
    For i = 1 To Len(invalidChars)
        result = Replace(result, Mid(invalidChars, i, 1), "_")
    Next i
    
    ' Trim whitespace
    result = Trim(result)
    
    ' Limit length to 200 characters (safety margin)
    If Len(result) > 200 Then
        result = Left(result, 200)
    End If
    
    SanitizeFilename = result
End Function

'-----------------------------------------------------------------------------
' GetCellValue
' Purpose: Safely retrieve cell value with error handling
' Parameters:
'   ws - Worksheet object
'   rowNum - Row number
'   colRef - Column reference (letter like "A" or number like 1)
' Returns: Cell value as string, empty string on error
'-----------------------------------------------------------------------------
Public Function GetCellValue(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal colRef As Variant) As String
    On Error Resume Next
    
    Dim value As Variant
    
    ' Handle column reference as letter or number
    If IsNumeric(colRef) Then
        value = ws.Cells(rowNum, CLng(colRef)).value
    Else
        value = ws.Range(colRef & rowNum).value
    End If
    
    ' Convert to string and trim
    If IsNull(value) Or IsEmpty(value) Then
        GetCellValue = ""
    Else
        GetCellValue = Trim(CStr(value))
    End If
    
    On Error GoTo 0
End Function

'-----------------------------------------------------------------------------
' BuildFormattedName
' Purpose: Construct full name from components, handling missing middle name
' Parameters:
'   firstName - First name
'   middleName - Middle name (optional)
'   lastName - Last name
' Returns: Formatted full name
' Examples:
'   BuildFormattedName("John", "Q", "Smith") -> "John Q Smith"
'   BuildFormattedName("John", "", "Smith") -> "John Smith"
'-----------------------------------------------------------------------------
Public Function BuildFormattedName(ByVal firstName As String, ByVal middleName As String, ByVal lastName As String) As String
    Dim result As String
    
    firstName = Trim(firstName)
    middleName = Trim(middleName)
    lastName = Trim(lastName)
    
    ' Build name with middle name if present
    If middleName <> "" Then
        result = firstName & " " & middleName & " " & lastName
    Else
        result = firstName & " " & lastName
    End If
    
    ' Clean up any double spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    BuildFormattedName = Trim(result)
End Function

'-----------------------------------------------------------------------------
' GetSettingValue
' Purpose: Retrieve setting from Settings sheet with fallback default
' Parameters:
'   settingName - Name of setting (matches cell in column A)
'   defaultValue - Value to return if setting not found or empty
' Returns: Setting value or default
'-----------------------------------------------------------------------------
Public Function GetSettingValue(ByVal settingName As String, ByVal defaultValue As String) As String
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    
    Set ws = ThisWorkbook.Sheets("Settings")
    If ws Is Nothing Then
        GetSettingValue = defaultValue
        Exit Function
    End If
    
    ' Find setting in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 1 To lastRow
        If Trim(LCase(ws.Cells(i, "A").value)) = Trim(LCase(settingName)) Then
            cellValue = Trim(ws.Cells(i, "B").value)
            
            ' Return setting if not empty, otherwise default
            If cellValue <> "" Then
                GetSettingValue = cellValue
            Else
                GetSettingValue = defaultValue
            End If
            Exit Function
        End If
    Next i
    
    ' Setting not found, return default
    GetSettingValue = defaultValue
    
    On Error GoTo 0
End Function

'-----------------------------------------------------------------------------
' DebugLog
' Purpose: Centralized debug logging to Immediate Window
' Parameters:
'   moduleName - Name of calling module
'   functionName - Name of calling function
'   message - Log message
' Note: Only logs if running in VBA editor (development mode)
'-----------------------------------------------------------------------------
Public Sub DebugLog(ByVal moduleName As String, ByVal functionName As String, ByVal message As String)
    ' Only log in development mode (VBA editor open)
    #If VBA7 Then
        Dim timestamp As String
        timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
        Debug.Print timestamp & " | " & moduleName & "." & functionName & " | " & message
    #End If
End Sub

'-----------------------------------------------------------------------------
' FormatPhoneNumber
' Purpose: Format phone number to (XXX) XXX-XXXX
' Parameters:
'   phoneNumber - Raw phone number string
' Returns: Formatted phone number
'-----------------------------------------------------------------------------
Public Function FormatPhoneNumber(ByVal phoneNumber As String) As String
    Dim digitsOnly As String
    Dim i As Integer
    Dim char As String
    
    ' Extract digits only
    digitsOnly = ""
    For i = 1 To Len(phoneNumber)
        char = Mid(phoneNumber, i, 1)
        If IsNumeric(char) Then
            digitsOnly = digitsOnly & char
        End If
    Next i
    
    ' Format if we have 10 digits
    If Len(digitsOnly) = 10 Then
        FormatPhoneNumber = "(" & Left(digitsOnly, 3) & ") " & _
                           Mid(digitsOnly, 4, 3) & "-" & _
                           Right(digitsOnly, 4)
    ElseIf Len(digitsOnly) = 11 And Left(digitsOnly, 1) = "1" Then
        ' Handle +1 country code
        FormatPhoneNumber = "(" & Mid(digitsOnly, 2, 3) & ") " & _
                           Mid(digitsOnly, 5, 3) & "-" & _
                           Right(digitsOnly, 4)
    Else
        ' Return as-is if not standard format
        FormatPhoneNumber = phoneNumber
    End If
End Function

'-----------------------------------------------------------------------------
' ValidateEmail
' Purpose: Basic email format validation
' Parameters:
'   email - Email address to validate
' Returns: True if valid format, False otherwise
'-----------------------------------------------------------------------------
Public Function ValidateEmail(ByVal email As String) As Boolean
    Dim atPos As Integer
    Dim dotPos As Integer
    
    email = Trim(email)
    
    ' Must contain @ and .
    atPos = InStr(email, "@")
    dotPos = InStrRev(email, ".")
    
    ' Basic validation
    If atPos > 1 And dotPos > atPos + 1 And dotPos < Len(email) Then
        ValidateEmail = True
    Else
        ValidateEmail = False
    End If
End Function

'-----------------------------------------------------------------------------
' EscapeHTML
' Purpose: Escape HTML special characters
' Parameters:
'   text - Text to escape
' Returns: HTML-safe text
'-----------------------------------------------------------------------------
Public Function EscapeHTML(ByVal text As String) As String
    Dim result As String
    
    result = text
    result = Replace(result, "&", "&amp;")
    result = Replace(result, "<", "&lt;")
    result = Replace(result, ">", "&gt;")
    result = Replace(result, """", "&quot;")
    result = Replace(result, "'", "&#39;")
    
    EscapeHTML = result
End Function
