Attribute VB_Name = "Module4_QRCode"
'=============================================================================
' Module4_QRCode.bas
' QR Code generation for QRVCARD Generator V4
'
' Purpose: Generate QR codes using Google Charts API
'
' Public Functions:
'   - GenerateQRCodeDataURL() - Creates QR code data URL for embedding in HTML
'
' Technical Details:
'   - Uses Google Charts API (https://chart.googleapis.com/chart)
'   - Returns data URL suitable for HTML <img src="">
'   - Requires internet connection
'   - Configurable size and error correction
'
' Dependencies:
'   - Module5_Utilities (for DebugLog)
'=============================================================================

Option Explicit

'-----------------------------------------------------------------------------
' GenerateQRCodeDataURL
' Purpose: Generate QR code image URL from vCard data
' Parameters:
'   vCardData - Complete vCard text to encode
'   sizePixels - QR code size in pixels (default: 250)
' Returns: Google Charts API URL for QR code image
'-----------------------------------------------------------------------------
Public Function GenerateQRCodeDataURL(ByVal vCardData As String, Optional ByVal sizePixels As Integer = 250) As String
    On Error GoTo ErrorHandler
    
    Dim encodedData As String
    Dim apiURL As String
    Dim errorCorrection As String
    Dim margin As Integer
    
    ' URL encode the vCard data
    encodedData = URLEncode(vCardData)
    
    ' QR code parameters
    errorCorrection = "M"  ' Error correction level: L, M, Q, H (M = medium, 15% recovery)
    margin = 1             ' Border size (modules)
    
    ' Build Google Charts API URL
    apiURL = "https://chart.googleapis.com/chart?" & _
             "cht=qr" & _
             "&chs=" & sizePixels & "x" & sizePixels & _
             "&chl=" & encodedData & _
             "&choe=UTF-8" & _
             "&chld=" & errorCorrection & "|" & margin
    
    GenerateQRCodeDataURL = apiURL
    
    Module5_Utilities.DebugLog "Module4_QRCode", "GenerateQRCodeDataURL", "QR URL generated, length: " & Len(apiURL)
    
    Exit Function
    
ErrorHandler:
    Module5_Utilities.DebugLog "Module4_QRCode", "GenerateQRCodeDataURL", "ERROR: " & Err.Description
    GenerateQRCodeDataURL = ""
End Function

'-----------------------------------------------------------------------------
' URLEncode
' Purpose: URL-encode text for use in query parameters
' Parameters:
'   text - Text to encode
' Returns: URL-encoded string
'-----------------------------------------------------------------------------
Private Function URLEncode(ByVal text As String) As String
    Dim i As Long
    Dim char As String
    Dim asciiValue As Integer
    Dim result As String
    
    result = ""
    
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        asciiValue = Asc(char)
        
        ' Check if character needs encoding
        If (asciiValue >= 48 And asciiValue <= 57) Or _
           (asciiValue >= 65 And asciiValue <= 90) Or _
           (asciiValue >= 97 And asciiValue <= 122) Or _
           char = "-" Or char = "_" Or char = "." Or char = "~" Then
            ' Safe character, no encoding needed
            result = result & char
        Else
            ' Encode character
            Select Case char
                Case " "
                    result = result & "%20"
                Case vbCr
                    result = result & "%0D"
                Case vbLf
                    result = result & "%0A"
                Case vbTab
                    result = result & "%09"
                Case ":"
                    result = result & "%3A"
                Case ";"
                    result = result & "%3B"
                Case ","
                    result = result & "%2C"
                Case "="
                    result = result & "%3D"
                Case "+"
                    result = result & "%2B"
                Case Else
                    ' Generic hex encoding
                    If asciiValue < 16 Then
                        result = result & "%0" & Hex(asciiValue)
                    Else
                        result = result & "%" & Hex(asciiValue)
                    End If
            End Select
        End If
    Next i
    
    URLEncode = result
End Function

'-----------------------------------------------------------------------------
' GetQRCodeImageSize
' Purpose: Retrieve QR code size from Settings sheet
' Returns: QR code size in pixels (default: 250)
'-----------------------------------------------------------------------------
Public Function GetQRCodeImageSize() As Integer
    Dim sizeValue As String
    Dim size As Integer
    
    sizeValue = Module5_Utilities.GetSettingValue("QR Code Size (pixels)", "250")
    
    ' Validate size
    If IsNumeric(sizeValue) Then
        size = CInt(sizeValue)
        
        ' Ensure reasonable size (100-500 pixels)
        If size < 100 Then size = 100
        If size > 500 Then size = 500
        
        GetQRCodeImageSize = size
    Else
        GetQRCodeImageSize = 250
    End If
End Function
