VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "Export Progress"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   OleObjectBlob   =   "frmProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================
' frmProgress.frm
' Progress bar UserForm for QRVCARD Generator V4
'
' Purpose: Display visual progress feedback during batch exports
'
' Controls Required (Create manually in VBA Editor):
'   - lblTitle        Label, top center: "Exporting Digital Business Cards"
'   - lblProgress     Label, middle: "Processing: [Name]"
'   - lblStatus       Label, below progress: "X of Y complete (Z%)"
'   - frmProgressBar  Frame, for progress bar background
'   - lblBar          Label inside frmProgressBar, for progress indicator
'
' Public Methods:
'   - InitializeProgress(totalCount)      Set up progress tracking
'   - UpdateProgress(itemNumber, itemName)  Update display
'
' CRITICAL: This form is MANDATORY and NON-OPTIONAL for batch exports
'
' Dependencies: None (standalone UI)
'=============================================================================

Option Explicit

Private mTotalCount As Long
Private mCurrentItem As Long

'-----------------------------------------------------------------------------
' InitializeProgress
' Purpose: Set up progress bar for batch operation
' Parameters:
'   totalCount - Total number of items to process
'-----------------------------------------------------------------------------
Public Sub InitializeProgress(ByVal totalCount As Long)
    mTotalCount = totalCount
    mCurrentItem = 0
    
    ' Set title
    Me.lblTitle.Caption = "Exporting Digital Business Cards"
    
    ' Initialize progress label
    Me.lblProgress.Caption = "Starting export..."
    
    ' Initialize status
    Me.lblStatus.Caption = "0 of " & totalCount & " complete (0%)"
    
    ' Initialize progress bar (0% width)
    Me.lblBar.Width = 0
    
    ' Refresh form
    Me.Repaint
    DoEvents
End Sub

'-----------------------------------------------------------------------------
' UpdateProgress
' Purpose: Update progress display for current item
' Parameters:
'   itemNumber - Current item number (1-based)
'   itemName - Name of item being processed
'-----------------------------------------------------------------------------
Public Sub UpdateProgress(ByVal itemNumber As Long, ByVal itemName As String)
    Dim percentComplete As Double
    Dim maxBarWidth As Double
    Dim currentBarWidth As Double
    
    mCurrentItem = itemNumber
    
    ' Calculate percentage
    If mTotalCount > 0 Then
        percentComplete = (itemNumber / mTotalCount) * 100
    Else
        percentComplete = 0
    End If
    
    ' Update progress label
    Me.lblProgress.Caption = "Processing: " & itemName
    
    ' Update status label
    Me.lblStatus.Caption = itemNumber & " of " & mTotalCount & " complete (" & Format(percentComplete, "0") & "%)"
    
    ' Update progress bar width
    maxBarWidth = Me.frmProgressBar.Width - 4  ' Leave 2px margin on each side
    currentBarWidth = maxBarWidth * (percentComplete / 100)
    Me.lblBar.Width = currentBarWidth
    
    ' Refresh form
    Me.Repaint
    DoEvents
End Sub

'-----------------------------------------------------------------------------
' UserForm_Initialize
' Purpose: Set up form when created
'-----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    ' Set form properties
    Me.Caption = "Export Progress"
    
    ' Initialize controls (if not already done in designer)
    ' These will be set up when controls are manually created
End Sub


'=============================================================================
' IMPORTANT SETUP INSTRUCTIONS
'=============================================================================
'
' This UserForm requires manual creation of controls in the VBA Editor.
' After importing this .frm file, follow these steps:
'
' 1. In VBA Editor, double-click frmProgress in the Project Explorer
' 2. If controls don't exist, add them from the Toolbox:
'
'    CONTROL 1: lblTitle (Label)
'    ---------------------------
'    - Name: lblTitle
'    - Caption: "Exporting Digital Business Cards"
'    - Font: Bold, 11pt
'    - TextAlign: Center (fmTextAlignCenter = 2)
'    - Top: 20
'    - Left: 50
'    - Width: 500
'    - Height: 20
'    - ForeColor: &H00000000 (Black)
'
'    CONTROL 2: lblProgress (Label)
'    ------------------------------
'    - Name: lblProgress
'    - Caption: "Processing: [Name]"
'    - Font: Regular, 9pt
'    - TextAlign: Left (fmTextAlignLeft = 1)
'    - Top: 60
'    - Left: 50
'    - Width: 500
'    - Height: 20
'    - ForeColor: &H00000000 (Black)
'
'    CONTROL 3: frmProgressBar (Frame)
'    ----------------------------------
'    - Name: frmProgressBar
'    - Caption: "" (empty)
'    - BackColor: &H00E0E0E0 (Light Gray)
'    - Top: 100
'    - Left: 50
'    - Width: 500
'    - Height: 30
'
'    CONTROL 4: lblBar (Label) - INSIDE frmProgressBar
'    --------------------------------------------------
'    - Name: lblBar
'    - Caption: "" (empty)
'    - BackColor: &H00FF8000 (Blue - #0066CC)
'    - Top: 2
'    - Left: 2
'    - Width: 0 (starts at 0, updated by code)
'    - Height: 26
'
'    CONTROL 5: lblStatus (Label)
'    ----------------------------
'    - Name: lblStatus
'    - Caption: "0 of 0 complete (0%)"
'    - Font: Regular, 9pt
'    - TextAlign: Center (fmTextAlignCenter = 2)
'    - Top: 150
'    - Left: 50
'    - Width: 500
'    - Height: 20
'    - ForeColor: &H00808080 (Gray)
'
' 3. Set Form Properties:
'    - Width: 600
'    - Height: 240
'    - StartUpPosition: 1 - CenterOwner
'    - ShowModal: False
'
' 4. Save the form
'
' The .frx file will be automatically created when you save.
'
'=============================================================================
