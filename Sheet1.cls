VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'=============================================================================
' Sheet1.cls (Data Entry sheet)
' Sheet-level event handlers for QRVCARD Generator V4
'
' Purpose: Handle data validation and auto-formatting
'
' Events:
'   - Worksheet_Change()   Triggered when cells change
'
' Validation Features:
'   - Name fields (B-D): Letters and spaces only
'   - Email (F): Valid email format
'   - Phone (G-H): Auto-format to (XXX) XXX-XXXX
'   - Extension (I): Numeric only
'
' Dependencies:
'   - Module5_Utilities (FormatPhoneNumber, ValidateEmail)
'=============================================================================

Option Explicit

'-----------------------------------------------------------------------------
' Worksheet_Change
' Purpose: Validate and format data as user types
' Parameters:
'   Target - Range that was changed
'-----------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    
    Dim cell As Range
    Dim ws As Worksheet
    Set ws = Me
    
    ' Disable events to prevent recursion
    Application.EnableEvents = False
    
    ' Process each changed cell
    For Each cell In Target
        If cell.Row >= 2 Then  ' Skip header row
            
            ' === COLUMN B-D: NAME VALIDATION (Letters and spaces only) ===
            If Not Intersect(cell, ws.Range("B:D")) Is Nothing Then
                ValidateNameField cell
            End If
            
            ' === COLUMN F: EMAIL VALIDATION ===
            If Not Intersect(cell, ws.Range("F:F")) Is Nothing Then
                ValidateEmailField cell
            End If
            
            ' === COLUMN G-H: PHONE AUTO-FORMATTING ===
            If Not Intersect(cell, ws.Range("G:H")) Is Nothing Then
                FormatPhoneField cell
            End If
            
            ' === COLUMN I: EXTENSION (Numeric only) ===
            If Not Intersect(cell, ws.Range("I:I")) Is Nothing Then
                ValidateExtensionField cell
            End If
            
        End If
    Next cell
    
    ' Re-enable events
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

'-----------------------------------------------------------------------------
' ValidateNameField
' Purpose: Validate name contains only letters and spaces
' Parameters:
'   cell - Cell to validate
'-----------------------------------------------------------------------------
Private Sub ValidateNameField(ByVal cell As Range)
    Dim value As String
    Dim i As Integer
    Dim char As String
    Dim isValid As Boolean
    
    value = Trim(cell.value)
    
    If value = "" Then
        cell.Interior.ColorIndex = xlNone
        Exit Sub
    End If
    
    isValid = True
    
    ' Check each character
    For i = 1 To Len(value)
        char = Mid(value, i, 1)
        
        ' Allow letters, spaces, hyphens, apostrophes
        If Not ((char >= "A" And char <= "Z") Or _
                (char >= "a" And char <= "z") Or _
                char = " " Or char = "-" Or char = "'") Then
            isValid = False
            Exit For
        End If
    Next i
    
    ' Highlight invalid entries
    If Not isValid Then
        cell.Interior.Color = RGB(255, 200, 200)  ' Light red
    Else
        cell.Interior.ColorIndex = xlNone
    End If
End Sub

'-----------------------------------------------------------------------------
' ValidateEmailField
' Purpose: Validate email format
' Parameters:
'   cell - Cell to validate
'-----------------------------------------------------------------------------
Private Sub ValidateEmailField(ByVal cell As Range)
    Dim email As String
    
    email = Trim(cell.value)
    
    If email = "" Then
        cell.Interior.ColorIndex = xlNone
        Exit Sub
    End If
    
    ' Use Module5_Utilities.ValidateEmail function
    If Module5_Utilities.ValidateEmail(email) Then
        cell.Interior.ColorIndex = xlNone
    Else
        cell.Interior.Color = RGB(255, 200, 200)  ' Light red
    End If
End Sub

'-----------------------------------------------------------------------------
' FormatPhoneField
' Purpose: Auto-format phone number to (XXX) XXX-XXXX
' Parameters:
'   cell - Cell to format
'-----------------------------------------------------------------------------
Private Sub FormatPhoneField(ByVal cell As Range)
    Dim phoneNumber As String
    Dim formatted As String
    
    phoneNumber = Trim(cell.value)
    
    If phoneNumber = "" Then Exit Sub
    
    ' Use Module5_Utilities.FormatPhoneNumber function
    formatted = Module5_Utilities.FormatPhoneNumber(phoneNumber)
    
    ' Only update if format changed
    If formatted <> phoneNumber Then
        cell.value = formatted
    End If
End Sub

'-----------------------------------------------------------------------------
' ValidateExtensionField
' Purpose: Validate extension is numeric only
' Parameters:
'   cell - Cell to validate
'-----------------------------------------------------------------------------
Private Sub ValidateExtensionField(ByVal cell As Range)
    Dim value As String
    
    value = Trim(cell.value)
    
    If value = "" Then
        cell.Interior.ColorIndex = xlNone
        Exit Sub
    End If
    
    ' Check if numeric
    If IsNumeric(value) Then
        cell.Interior.ColorIndex = xlNone
        
        ' Format as text to prevent Excel from adding decimals
        cell.NumberFormat = "@"
    Else
        cell.Interior.Color = RGB(255, 200, 200)  ' Light red
    End If
End Sub
