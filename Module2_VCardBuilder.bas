Attribute VB_Name = "Module2_VCardBuilder"
'=============================================================================
' Module2_VCardBuilder.bas
' vCard 3.0 generation for QRVCARD Generator V4
'
' Purpose: Generate RFC 2426 compliant vCard files
'
' Public Functions:
'   - SaveVcard()         Generate and save vCard file
'   - GenerateVcard()     Create vCard 3.0 formatted string
'
' vCard 3.0 Specification: RFC 2426
' Character encoding: UTF-8
'
' Dependencies:
'   - Module5_Utilities (for GetCellValue, BuildFormattedName, DebugLog)
'=============================================================================

Option Explicit

'-----------------------------------------------------------------------------
' SaveVcard
' Purpose: Generate vCard and save to file
' Parameters:
'   ws - Data Entry worksheet
'   rowNum - Row number to export
'   filePath - Full path for output file
' Returns: True on success, False on failure
'-----------------------------------------------------------------------------
Public Function SaveVcard(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim vCardContent As String
    Dim fileNum As Integer
    
    ' Generate vCard content
    vCardContent = GenerateVcard(ws, rowNum)
    
    If vCardContent = "" Then
        Module5_Utilities.DebugLog "Module2_VCard", "SaveVcard", "ERROR: Empty vCard content"
        SaveVcard = False
        Exit Function
    End If
    
    ' Write to file
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, vCardContent
    Close #fileNum
    
    Module5_Utilities.DebugLog "Module2_VCard", "SaveVcard", "SUCCESS: " & filePath
    SaveVcard = True
    
    Exit Function
    
ErrorHandler:
    Module5_Utilities.DebugLog "Module2_VCard", "SaveVcard", "ERROR: " & Err.Description
    SaveVcard = False
End Function

'-----------------------------------------------------------------------------
' GenerateVcard
' Purpose: Create vCard 3.0 formatted string from row data
' Parameters:
'   ws - Data Entry worksheet
'   rowNum - Row number to export
' Returns: vCard 3.0 text
'-----------------------------------------------------------------------------
Public Function GenerateVcard(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    On Error GoTo ErrorHandler
    
    Dim vCard As String
    Dim lastName As String, firstName As String, middleName As String
    Dim fullName As String
    Dim email As String
    Dim cellPhone As String, officePhone As String, extension As String
    Dim titlePrimary As String, titleSecondary As String
    Dim orgPrimary As String, orgSecondary As String
    Dim street As String, city As String, state As String, postalCode As String, country As String
    Dim website As String
    Dim bio As String
    Dim noteText As String
    
    ' Extract data from columns using GetCellValue (39-column structure)
    lastName = Module5_Utilities.GetCellValue(ws, rowNum, "B")      ' B: Last Name
    firstName = Module5_Utilities.GetCellValue(ws, rowNum, "C")     ' C: First Name
    middleName = Module5_Utilities.GetCellValue(ws, rowNum, "D")    ' D: Middle Name
    email = Module5_Utilities.GetCellValue(ws, rowNum, "F")         ' F: Email
    cellPhone = Module5_Utilities.GetCellValue(ws, rowNum, "G")     ' G: Phone Cell
    officePhone = Module5_Utilities.GetCellValue(ws, rowNum, "H")   ' H: Phone Office
    extension = Module5_Utilities.GetCellValue(ws, rowNum, "I")     ' I: Extension
    titlePrimary = Module5_Utilities.GetCellValue(ws, rowNum, "J")  ' J: Title Primary
    titleSecondary = Module5_Utilities.GetCellValue(ws, rowNum, "K") ' K: Title Secondary
    orgPrimary = Module5_Utilities.GetCellValue(ws, rowNum, "L")    ' L: Organization Primary
    orgSecondary = Module5_Utilities.GetCellValue(ws, rowNum, "M")  ' M: Organization Secondary
    street = Module5_Utilities.GetCellValue(ws, rowNum, "N")        ' N: Street Address
    city = Module5_Utilities.GetCellValue(ws, rowNum, "O")          ' O: City
    state = Module5_Utilities.GetCellValue(ws, rowNum, "P")         ' P: State
    postalCode = Module5_Utilities.GetCellValue(ws, rowNum, "Q")    ' Q: Postal Code
    country = Module5_Utilities.GetCellValue(ws, rowNum, "R")       ' R: Country
    website = Module5_Utilities.GetCellValue(ws, rowNum, "S")       ' S: Website URL
    bio = Module5_Utilities.GetCellValue(ws, rowNum, "AC")          ' AC: Professional Bio
    
    ' Build full name
    fullName = Module5_Utilities.BuildFormattedName(firstName, middleName, lastName)
    
    ' Start vCard 3.0
    vCard = "BEGIN:VCARD" & vbCrLf
    vCard = vCard & "VERSION:3.0" & vbCrLf
    
    ' N: Structured name (Last;First;Middle;Prefix;Suffix)
    vCard = vCard & "N:" & lastName & ";" & firstName
    If middleName <> "" Then
        vCard = vCard & ";" & middleName
    End If
    vCard = vCard & ";;" & vbCrLf
    
    ' FN: Formatted name (display name)
    vCard = vCard & "FN:" & fullName & vbCrLf
    
    ' ORG: Organization (Primary)
    If orgPrimary <> "" Then
        vCard = vCard & "ORG:" & orgPrimary & vbCrLf
    End If
    
    ' TITLE: Job title (Primary)
    If titlePrimary <> "" Then
        vCard = vCard & "TITLE:" & titlePrimary & vbCrLf
    End If
    
    ' TEL: Phone numbers (Cell is primary, Office is secondary)
    If cellPhone <> "" Then
        vCard = vCard & "TEL;TYPE=CELL,VOICE:" & cellPhone & vbCrLf
    End If
    
    If officePhone <> "" Then
        If extension <> "" Then
            vCard = vCard & "TEL;TYPE=WORK,VOICE:" & officePhone & " x" & extension & vbCrLf
        Else
            vCard = vCard & "TEL;TYPE=WORK,VOICE:" & officePhone & vbCrLf
        End If
    End If
    
    ' EMAIL: Email address
    If email <> "" Then
        vCard = vCard & "EMAIL;TYPE=INTERNET,PREF:" & email & vbCrLf
    End If
    
    ' URL: Website
    If website <> "" Then
        vCard = vCard & "URL:" & website & vbCrLf
    End If
    
    ' ADR: Address (;;Street;City;State;Postal;Country)
    If street <> "" Or city <> "" Or state <> "" Or postalCode <> "" Then
        vCard = vCard & "ADR;TYPE=WORK:;;" & street & ";" & city & ";" & state & ";" & postalCode & ";" & country & vbCrLf
    End If
    
    ' NOTE: Build comprehensive note with secondary role and bio
    noteText = ""
    
    ' Add secondary role if present
    If titleSecondary <> "" Or orgSecondary <> "" Then
        If titleSecondary <> "" And orgSecondary <> "" Then
            noteText = "Also serves as " & titleSecondary & " at " & orgSecondary & "."
        ElseIf titleSecondary <> "" Then
            noteText = "Also serves as " & titleSecondary & "."
        ElseIf orgSecondary <> "" Then
            noteText = "Also with " & orgSecondary & "."
        End If
    End If
    
    ' Add bio if present
    If bio <> "" Then
        If noteText <> "" Then
            noteText = noteText & " " & bio
        Else
            noteText = bio
        End If
    End If
    
    ' Write NOTE field if we have content
    If noteText <> "" Then
        vCard = vCard & "NOTE:" & noteText & vbCrLf
    End If
    
    ' REV: Revision timestamp
    vCard = vCard & "REV:" & Format(Now, "yyyy-mm-ddThh:nn:ssZ") & vbCrLf
    
    ' End vCard
    vCard = vCard & "END:VCARD"
    
    GenerateVcard = vCard
    
    Module5_Utilities.DebugLog "Module2_VCard", "GenerateVcard", "vCard generated for: " & fullName
    
    Exit Function
    
ErrorHandler:
    Module5_Utilities.DebugLog "Module2_VCard", "GenerateVcard", "ERROR: " & Err.Description
    GenerateVcard = ""
End Function
