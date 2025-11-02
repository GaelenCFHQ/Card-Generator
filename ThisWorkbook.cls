VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'=============================================================================
' ThisWorkbook.cls
' Workbook-level event handlers for QRVCARD Generator V4
'
' Purpose: Handle workbook open/close events
'
' Events:
'   - Workbook_Open()     Run on workbook open
'   - Workbook_BeforeClose()   Run before workbook close
'
' Dependencies: None
'=============================================================================

Option Explicit

'-----------------------------------------------------------------------------
' Workbook_Open
' Purpose: Initialize workbook on open
'-----------------------------------------------------------------------------
Private Sub Workbook_Open()
    ' Display welcome message (optional)
    ' MsgBox "QRVCARD Generator V4 loaded successfully!", vbInformation, "Welcome"
    
    ' Set active sheet to Data Entry
    On Error Resume Next
    ThisWorkbook.Sheets("Data Entry").Activate
    On Error GoTo 0
    
    ' Log workbook open
    Module5_Utilities.DebugLog "ThisWorkbook", "Workbook_Open", "Workbook opened"
End Sub

'-----------------------------------------------------------------------------
' Workbook_BeforeClose
' Purpose: Cleanup before workbook closes
' Parameters:
'   Cancel - Set to True to cancel close operation
'-----------------------------------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Log workbook close
    Module5_Utilities.DebugLog "ThisWorkbook", "Workbook_BeforeClose", "Workbook closing"
    
    ' Optional: Prompt to save if unsaved changes
    ' If Not ThisWorkbook.Saved Then
    '     Dim response As VbMsgBoxResult
    '     response = MsgBox("Save changes before closing?", vbYesNoCancel + vbQuestion, "Save Changes")
    '     
    '     If response = vbYes Then
    '         ThisWorkbook.Save
    '     ElseIf response = vbCancel Then
    '         Cancel = True
    '     End If
    ' End If
End Sub
