Attribute VB_Name = "Module3_HTMLBuilder"
'=============================================================================
' Module3_HTMLBuilder.bas
' HTML landing page generation for QRVCARD Generator V4
'
' Purpose: Generate responsive HTML contact cards with social media icons
'
' Public Functions:
'   - SaveHTML()              Generate and save HTML file
'
' Features:
'   - Responsive modern design
'   - Social media icon integration (10 platforms)
'   - Embedded QR code
'   - Dual role display
'   - Google Analytics tracking (optional)
'   - Theme support
'
' Dependencies:
'   - Module2_VCardBuilder (for vCard content)
'   - Module4_QRCode (for QR code URL)
'   - Module5_Utilities (for helper functions)
'=============================================================================

Option Explicit

'-----------------------------------------------------------------------------
' SaveHTML
' Purpose: Generate HTML landing page and save to file
' Parameters:
'   ws - Data Entry worksheet
'   wsSettings - Settings worksheet
'   rowNum - Row number to export
'   filePath - Full path for output file
' Returns: True on success, False on failure
'-----------------------------------------------------------------------------
Public Function SaveHTML(ByVal ws As Worksheet, ByVal wsSettings As Worksheet, ByVal rowNum As Long, ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim htmlContent As String
    Dim fileNum As Integer
    
    ' Generate HTML content
    htmlContent = GenerateHTML(ws, wsSettings, rowNum)
    
    If htmlContent = "" Then
        Module5_Utilities.DebugLog "Module3_HTML", "SaveHTML", "ERROR: Empty HTML content"
        SaveHTML = False
        Exit Function
    End If
    
    ' Write to file with UTF-8 encoding
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, htmlContent
    Close #fileNum
    
    Module5_Utilities.DebugLog "Module3_HTML", "SaveHTML", "SUCCESS: " & filePath
    SaveHTML = True
    
    Exit Function
    
ErrorHandler:
    Module5_Utilities.DebugLog "Module3_HTML", "SaveHTML", "ERROR: " & Err.Description
    SaveHTML = False
End Function

'-----------------------------------------------------------------------------
' GenerateHTML
' Purpose: Create complete HTML document
' Parameters:
'   ws - Data Entry worksheet
'   wsSettings - Settings worksheet
'   rowNum - Row number to export
' Returns: Complete HTML document as string
'-----------------------------------------------------------------------------
Private Function GenerateHTML(ByVal ws As Worksheet, ByVal wsSettings As Worksheet, ByVal rowNum As Long) As String
    On Error GoTo ErrorHandler
    
    Dim html As String
    Dim fullName As String, firstName As String, lastName As String
    Dim titlePrimary As String, titleSecondary As String
    Dim orgPrimary As String, orgSecondary As String
    Dim email As String, cellPhone As String, officePhone As String, extension As String
    Dim website As String, bio As String
    Dim vCardContent As String, qrCodeURL As String
    Dim theme As String, headerColor As String
    Dim analyticsEnabled As String, franchiseCode As String
    
    ' Extract basic data
    firstName = Module5_Utilities.GetCellValue(ws, rowNum, "C")
    lastName = Module5_Utilities.GetCellValue(ws, rowNum, "B")
    fullName = Module5_Utilities.GetCellValue(ws, rowNum, "E")
    email = Module5_Utilities.GetCellValue(ws, rowNum, "F")
    cellPhone = Module5_Utilities.GetCellValue(ws, rowNum, "G")
    officePhone = Module5_Utilities.GetCellValue(ws, rowNum, "H")
    extension = Module5_Utilities.GetCellValue(ws, rowNum, "I")
    titlePrimary = Module5_Utilities.GetCellValue(ws, rowNum, "J")
    titleSecondary = Module5_Utilities.GetCellValue(ws, rowNum, "K")
    orgPrimary = Module5_Utilities.GetCellValue(ws, rowNum, "L")
    orgSecondary = Module5_Utilities.GetCellValue(ws, rowNum, "M")
    website = Module5_Utilities.GetCellValue(ws, rowNum, "S")
    bio = Module5_Utilities.GetCellValue(ws, rowNum, "AC")
    theme = Module5_Utilities.GetCellValue(ws, rowNum, "AF")
    franchiseCode = Module5_Utilities.GetCellValue(ws, rowNum, "AG")
    analyticsEnabled = Module5_Utilities.GetCellValue(ws, rowNum, "AH")
    
    ' Get theme color
    If theme = "Dark" Then
        headerColor = "#1a1a1a"
    ElseIf theme = "Corporate" Then
        headerColor = Module5_Utilities.GetSettingValue("Default Header Color (Hex)", "#0066CC")
    ElseIf theme = "Minimal" Then
        headerColor = "#666666"
    Else
        headerColor = "#0066CC"  ' Default theme
    End If
    
    ' Generate vCard content for QR code
    vCardContent = Module2_VCardBuilder.GenerateVcard(ws, rowNum)
    
    ' Generate QR code URL
    Dim qrSize As Integer
    qrSize = Module4_QRCode.GetQRCodeImageSize()
    qrCodeURL = Module4_QRCode.GenerateQRCodeDataURL(vCardContent, qrSize)
    
    ' Build HTML document
    html = "<!DOCTYPE html>" & vbCrLf
    html = html & "<html lang=""en"">" & vbCrLf
    html = html & "<head>" & vbCrLf
    html = html & "    <meta charset=""UTF-8"">" & vbCrLf
    html = html & "    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & vbCrLf
    html = html & "    <title>" & Module5_Utilities.EscapeHTML(fullName) & " - Digital Contact Card</title>" & vbCrLf
    html = html & GetCSS(headerColor)
    html = html & "</head>" & vbCrLf
    html = html & "<body>" & vbCrLf
    html = html & "    <div class=""container"">" & vbCrLf
    html = html & "        <div class=""card"">" & vbCrLf
    
    ' Header section
    html = html & "            <div class=""header"">" & vbCrLf
    html = html & "                <h1>" & Module5_Utilities.EscapeHTML(fullName) & "</h1>" & vbCrLf
    
    ' Display primary role
    If titlePrimary <> "" And orgPrimary <> "" Then
        html = html & "                <div class=""role"">" & Module5_Utilities.EscapeHTML(titlePrimary) & "</div>" & vbCrLf
        html = html & "                <div class=""org"">" & Module5_Utilities.EscapeHTML(orgPrimary) & "</div>" & vbCrLf
    ElseIf titlePrimary <> "" Then
        html = html & "                <div class=""role"">" & Module5_Utilities.EscapeHTML(titlePrimary) & "</div>" & vbCrLf
    End If
    
    ' Display secondary role if present
    If titleSecondary <> "" And orgSecondary <> "" Then
        html = html & "                <div class=""role-secondary"">" & Module5_Utilities.EscapeHTML(titleSecondary) & "</div>" & vbCrLf
        html = html & "                <div class=""org-secondary"">" & Module5_Utilities.EscapeHTML(orgSecondary) & "</div>" & vbCrLf
    ElseIf titleSecondary <> "" Then
        html = html & "                <div class=""role-secondary"">" & Module5_Utilities.EscapeHTML(titleSecondary) & "</div>" & vbCrLf
    End If
    
    html = html & "            </div>" & vbCrLf
    
    ' Content section
    html = html & "            <div class=""content"">" & vbCrLf
    
    ' Contact information
    If cellPhone <> "" Then
        html = html & "                <div class=""info-row"">" & vbCrLf
        html = html & "                    <div class=""info-label"">Mobile:</div>" & vbCrLf
        html = html & "                    <div class=""info-value""><a href=""tel:" & cellPhone & """>" & Module5_Utilities.EscapeHTML(cellPhone) & "</a></div>" & vbCrLf
        html = html & "                </div>" & vbCrLf
    End If
    
    If officePhone <> "" Then
        Dim phoneDisplay As String
        If extension <> "" Then
            phoneDisplay = officePhone & " x" & extension
        Else
            phoneDisplay = officePhone
        End If
        html = html & "                <div class=""info-row"">" & vbCrLf
        html = html & "                    <div class=""info-label"">Office:</div>" & vbCrLf
        html = html & "                    <div class=""info-value""><a href=""tel:" & officePhone & """>" & Module5_Utilities.EscapeHTML(phoneDisplay) & "</a></div>" & vbCrLf
        html = html & "                </div>" & vbCrLf
    End If
    
    If email <> "" Then
        html = html & "                <div class=""info-row"">" & vbCrLf
        html = html & "                    <div class=""info-label"">Email:</div>" & vbCrLf
        html = html & "                    <div class=""info-value""><a href=""mailto:" & email & """>" & Module5_Utilities.EscapeHTML(email) & "</a></div>" & vbCrLf
        html = html & "                </div>" & vbCrLf
    End If
    
    If website <> "" Then
        html = html & "                <div class=""info-row"">" & vbCrLf
        html = html & "                    <div class=""info-label"">Website:</div>" & vbCrLf
        html = html & "                    <div class=""info-value""><a href=""" & website & """ target=""_blank"">" & Module5_Utilities.EscapeHTML(website) & "</a></div>" & vbCrLf
        html = html & "                </div>" & vbCrLf
    End If
    
    ' Professional bio
    If bio <> "" Then
        html = html & "                <div class=""bio"">" & vbCrLf
        html = html & "                    <p>" & Module5_Utilities.EscapeHTML(bio) & "</p>" & vbCrLf
        html = html & "                </div>" & vbCrLf
    End If
    
    ' Social media icons
    html = html & GetSocialMediaIcons(ws, rowNum)
    
    ' QR Code section
    html = html & "                <div class=""qr-section"">" & vbCrLf
    html = html & "                    <img src=""" & qrCodeURL & """ alt=""Contact QR Code"" class=""qr-code"">" & vbCrLf
    html = html & "                    <div class=""qr-label"">Scan to add contact</div>" & vbCrLf
    html = html & "                </div>" & vbCrLf
    
    ' Download vCard button
    html = html & "                <div class=""download-section"">" & vbCrLf
    html = html & "                    <a href=""data:text/vcard;charset=utf-8," & URLEncode(vCardContent) & """ download=""" & lastName & "_" & firstName & ".vcf"" class=""download-btn"">Download Contact</a>" & vbCrLf
    html = html & "                </div>" & vbCrLf
    
    html = html & "            </div>" & vbCrLf
    html = html & "        </div>" & vbCrLf
    html = html & "    </div>" & vbCrLf
    
    ' Google Analytics (if enabled)
    If UCase(analyticsEnabled) = "TRUE" Then
        html = html & GetAnalyticsScript(wsSettings, fullName, titlePrimary, franchiseCode)
    End If
    
    html = html & "</body>" & vbCrLf
    html = html & "</html>"
    
    GenerateHTML = html
    
    Exit Function
    
ErrorHandler:
    Module5_Utilities.DebugLog "Module3_HTML", "GenerateHTML", "ERROR: " & Err.Description
    GenerateHTML = ""
End Function

'-----------------------------------------------------------------------------
' GetCSS
' Purpose: Generate CSS stylesheet
' Parameters:
'   headerColor - Header background color (hex code)
' Returns: CSS wrapped in <style> tags
'-----------------------------------------------------------------------------
Private Function GetCSS(ByVal headerColor As String) As String
    Dim css As String
    
    css = "    <style>" & vbCrLf
    css = css & "        * { margin: 0; padding: 0; box-sizing: border-box; }" & vbCrLf
    css = css & "        body {" & vbCrLf
    css = css & "            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;" & vbCrLf
    css = css & "            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);" & vbCrLf
    css = css & "            min-height: 100vh;" & vbCrLf
    css = css & "            display: flex;" & vbCrLf
    css = css & "            justify-content: center;" & vbCrLf
    css = css & "            align-items: center;" & vbCrLf
    css = css & "            padding: 20px;" & vbCrLf
    css = css & "        }" & vbCrLf
    css = css & "        .container { max-width: 500px; width: 100%; }" & vbCrLf
    css = css & "        .card {" & vbCrLf
    css = css & "            background: white;" & vbCrLf
    css = css & "            border-radius: 20px;" & vbCrLf
    css = css & "            box-shadow: 0 20px 60px rgba(0,0,0,0.3);" & vbCrLf
    css = css & "            overflow: hidden;" & vbCrLf
    css = css & "        }" & vbCrLf
    css = css & "        .header {" & vbCrLf
    css = css & "            background: " & headerColor & ";" & vbCrLf
    css = css & "            color: white;" & vbCrLf
    css = css & "            padding: 40px 30px;" & vbCrLf
    css = css & "            text-align: center;" & vbCrLf
    css = css & "        }" & vbCrLf
    css = css & "        .header h1 { font-size: 32px; margin-bottom: 10px; font-weight: 700; }" & vbCrLf
    css = css & "        .role { font-size: 18px; font-weight: 600; margin-top: 10px; }" & vbCrLf
    css = css & "        .org { font-size: 16px; opacity: 0.9; margin-top: 5px; }" & vbCrLf
    css = css & "        .role-secondary { font-size: 16px; font-weight: 500; margin-top: 15px; opacity: 0.85; }" & vbCrLf
    css = css & "        .org-secondary { font-size: 14px; opacity: 0.8; margin-top: 3px; }" & vbCrLf
    css = css & "        .content { padding: 30px; }" & vbCrLf
    css = css & "        .info-row {" & vbCrLf
    css = css & "            display: flex;" & vbCrLf
    css = css & "            margin-bottom: 15px;" & vbCrLf
    css = css & "            padding: 10px;" & vbCrLf
    css = css & "            border-radius: 8px;" & vbCrLf
    css = css & "            transition: background 0.2s;" & vbCrLf
    css = css & "        }" & vbCrLf
    css = css & "        .info-row:hover { background: #f8f9fa; }" & vbCrLf
    css = css & "        .info-label {" & vbCrLf
    css = css & "            font-weight: 600;" & vbCrLf
    css = css & "            color: #555;" & vbCrLf
    css = css & "            min-width: 90px;" & vbCrLf
    css = css & "            font-size: 14px;" & vbCrLf
    css = css & "        }" & vbCrLf
    css = css & "        .info-value {" & vbCrLf
    css = css & "            color: #333;" & vbCrLf
    css = css & "            font-size: 14px;" & vbCrLf
    css = css & "            word-break: break-word;" & vbCrLf
    css = css & "        }" & vbCrLf
    css = css & "        .info-value a {" & vbCrLf
    css = css & "            color: " & headerColor & ";" & vbCrLf
    css = css & "            text-decoration: none;" & vbCrLf
    css = css & "            border-bottom: 1px solid transparent;" & vbCrLf
    css = css & "            transition: border-color 0.2s;" & vbCrLf
    css = css & "        }" & vbCrLf
    css = css & "        .info-value a:hover { border-bottom-color: " & headerColor & "; }" & vbCrLf
    css = css & "        .bio {" & vbCrLf
    css = css & "            margin: 20px 0;" & vbCrLf
    css = css & "            padding: 15px;" & vbCrLf
    css = css & "            background: #f8f9fa;" & vbCrLf
    css = css & "            border-radius: 8px;" & vbCrLf
    css = css & "            line-height: 1.6;" & vbCrLf
    css = css & "            color: #555;" & vbCrLf
    css = css & "            font-size: 14px;" & vbCrLf
    css = css & "        }" & vbCrLf
    css = css & "        .social-icons {" & vbCrLf
    css = css & "            display: flex;" & vbCrLf
    css = css & "            justify-content: center;" & vbCrLf
    css = css & "            flex-wrap: wrap;" & vbCrLf
    css = css & "            gap: 12px;" & vbCrLf
    css = css & "            margin: 25px 0;" & vbCrLf
    css = css & "        }" & vbCrLf
    css = css & "        .social-icon {" & vbCrLf
    css = css & "            width: 44px;" & vbCrLf
    css = css & "            height: 44px;" & vbCrLf
    css = css & "            border-radius: 50%;" & vbCrLf
    css = css & "            display: flex;" & vbCrLf
    css = css & "            align-items: center;" & vbCrLf
    css = css & "            justify-content: center;" & vbCrLf
    css = css & "            font-size: 20px;" & vbCrLf
    css = css & "            color: white;" & vbCrLf
    css = css & "            text-decoration: none;" & vbCrLf
    css = css & "            transition: transform 0.2s, box-shadow 0.2s;" & vbCrLf
    css = css & "        }" & vbCrLf
    css = css & "        .social-icon:hover { transform: translateY(-3px); box-shadow: 0 4px 12px rgba(0,0,0,0.2); }" & vbCrLf
    css = css & "        .qr-section {" & vbCrLf
    css = css & "            text-align: center;" & vbCrLf
    css = css & "            margin: 30px 0;" & vbCrLf
    css = css & "        }" & vbCrLf
    css = css & "        .qr-code {" & vbCrLf
    css = css & "            max-width: 200px;" & vbCrLf
    css = css & "            width: 100%;" & vbCrLf
    css = css & "            height: auto;" & vbCrLf
    css = css & "            border-radius: 10px;" & vbCrLf
    css = css & "            box-shadow: 0 4px 12px rgba(0,0,0,0.1);" & vbCrLf
    css = css & "        }" & vbCrLf
    css = css & "        .qr-label {" & vbCrLf
    css = css & "            margin-top: 10px;" & vbCrLf
    css = css & "            font-size: 13px;" & vbCrLf
    css = css & "            color: #777;" & vbCrLf
    css = css & "        }" & vbCrLf
    css = css & "        .download-section { text-align: center; margin-top: 20px; }" & vbCrLf
    css = css & "        .download-btn {" & vbCrLf
    css = css & "            display: inline-block;" & vbCrLf
    css = css & "            padding: 12px 30px;" & vbCrLf
    css = css & "            background: " & headerColor & ";" & vbCrLf
    css = css & "            color: white;" & vbCrLf
    css = css & "            text-decoration: none;" & vbCrLf
    css = css & "            border-radius: 25px;" & vbCrLf
    css = css & "            font-weight: 600;" & vbCrLf
    css = css & "            transition: transform 0.2s, box-shadow 0.2s;" & vbCrLf
    css = css & "            box-shadow: 0 4px 12px rgba(0,0,0,0.15);" & vbCrLf
    css = css & "        }" & vbCrLf
    css = css & "        .download-btn:hover {" & vbCrLf
    css = css & "            transform: translateY(-2px);" & vbCrLf
    css = css & "            box-shadow: 0 6px 16px rgba(0,0,0,0.2);" & vbCrLf
    css = css & "        }" & vbCrLf
    css = css & "        @media (max-width: 600px) {" & vbCrLf
    css = css & "            .header h1 { font-size: 26px; }" & vbCrLf
    css = css & "            .role { font-size: 16px; }" & vbCrLf
    css = css & "            .content { padding: 20px; }" & vbCrLf
    css = css & "        }" & vbCrLf
    css = css & "    </style>" & vbCrLf
    
    GetCSS = css
End Function

'-----------------------------------------------------------------------------
' GetSocialMediaIcons
' Purpose: Generate social media icons HTML WITH INLINE SVG
' Parameters:
'   ws - Data Entry worksheet
'   rowNum - Row number
' Returns: HTML for social icons section with embedded SVG code
'-----------------------------------------------------------------------------
Private Function GetSocialMediaIcons(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    Dim html As String
    Dim url As String
    Dim iconAdded As Boolean
    
    iconAdded = False
    html = ""
    
    ' LinkedIn (Column T)
    url = Module5_Utilities.GetCellValue(ws, rowNum, "T")
    If url <> "" Then
        If Not iconAdded Then
            html = "                <div class=""social-icons"">" & vbCrLf
            iconAdded = True
        End If
        html = html & "                    <a href=""" & url & """ target=""_blank"" class=""social-icon"" title=""LinkedIn"">" & _
               GetLinkedInSVG() & "</a>" & vbCrLf
    End If
    
    ' Instagram (Column U)
    url = Module5_Utilities.GetCellValue(ws, rowNum, "U")
    If url <> "" Then
        If Not iconAdded Then
            html = "                <div class=""social-icons"">" & vbCrLf
            iconAdded = True
        End If
        html = html & "                    <a href=""" & url & """ target=""_blank"" class=""social-icon"" title=""Instagram"">" & _
               GetInstagramSVG() & "</a>" & vbCrLf
    End If
    
    ' Facebook (Column V)
    url = Module5_Utilities.GetCellValue(ws, rowNum, "V")
    If url <> "" Then
        If Not iconAdded Then
            html = "                <div class=""social-icons"">" & vbCrLf
            iconAdded = True
        End If
        html = html & "                    <a href=""" & url & """ target=""_blank"" class=""social-icon"" title=""Facebook"">" & _
               GetFacebookSVG() & "</a>" & vbCrLf
    End If
    
    ' Pinterest (Column X)
    url = Module5_Utilities.GetCellValue(ws, rowNum, "X")
    If url <> "" Then
        If Not iconAdded Then
            html = "                <div class=""social-icons"">" & vbCrLf
            iconAdded = True
        End If
        html = html & "                    <a href=""" & url & """ target=""_blank"" class=""social-icon"" title=""Pinterest"">" & _
               GetPinterestSVG() & "</a>" & vbCrLf
    End If
    
    ' Houzz (Column Y)
    url = Module5_Utilities.GetCellValue(ws, rowNum, "Y")
    If url <> "" Then
        If Not iconAdded Then
            html = "                <div class=""social-icons"">" & vbCrLf
            iconAdded = True
        End If
        html = html & "                    <a href=""" & url & """ target=""_blank"" class=""social-icon"" title=""Houzz"">" & _
               GetHouzzSVG() & "</a>" & vbCrLf
    End If
    
    ' Cal.com (Column AC)
    url = Module5_Utilities.GetCellValue(ws, rowNum, "AC")
    If url <> "" Then
        If Not iconAdded Then
            html = "                <div class=""social-icons"">" & vbCrLf
            iconAdded = True
        End If
        html = html & "                    <a href=""" & url & """ target=""_blank"" class=""social-icon"" title=""Schedule Meeting"">" & _
               GetCalComSVG() & "</a>" & vbCrLf
    End If
    
    ' Close social icons container
    If iconAdded Then
        html = html & "                </div>" & vbCrLf
    End If
    
    GetSocialMediaIcons = html
End Function

'=============================================================================
' INLINE SVG ICON FUNCTIONS - All using unified #284b66 color
'=============================================================================

Private Function GetLinkedInSVG() As String
    Dim p As String
    p = "M19 35h-4V21h4v14zm-2-16c-1.3 0-2.3-1-2.3-2.3S15.7 14.4 17 14.4s2.3 1 2.3 2.3-1 2.3-2.3 2.3z"
    p = p & "M35 35h-4v-7c0-1.7-.6-2.8-2-2.8-1.1 0-1.7.7-2 1.4-.1.3-.1.6-.1 1v7.4h-4s.1-12 0-13.2h4v1.9c.5-.8 1.4-2 3.5-2 "
    p = p & "2.5 0 4.4 1.7 4.4 5.2V35h.2z"
    GetLinkedInSVG = "<svg viewBox=""0 0 48 48""><circle fill=""#284b66"" cx=""24"" cy=""24"" r=""24""/>" & _
                     "<path fill=""#fff"" d=""" & p & """/></svg>"
End Function

Private Function GetInstagramSVG() As String
    Dim svg As String
    Dim p1 As String, p2 As String
    
    p1 = "M24 15.6c2.7 0 3 0 4.1.1 1 0 1.5.2 1.9.4.5.2.8.4 1.2.8.4.4.6.7.8 1.2.2.4.3.9.4 1.9.1 1.1.1 1.4.1 4.1s0 3-.1 4.1"
    p1 = p1 & "c0 1-.2 1.5-.4 1.9-.2.5-.4.8-.8 1.2-.4.4-.7.6-1.2.8-.4.2-.9.3-1.9.4-1.1.1-1.4.1-4.1.1s-3 0-4.1-.1c-1 0-1.5-.2-1.9-.4"
    p1 = p1 & "-.5-.2-.8-.4-1.2-.8-.4-.4-.6-.7-.8-1.2-.2-.4-.3-.9-.4-1.9-.1-1.1-.1-1.4-.1-4.1s0-3 .1-4.1c0-1 .2-1.5.4-1.9"
    p1 = p1 & ".2-.5.4-.8.8-1.2.4-.4.7-.6 1.2-.8.4-.2.9-.3 1.9-.4 1.1-.1 1.4-.1 4.1-.1m0-1.6c-2.7 0-3.1 0-4.2.1-1.1.1-1.8.2-2.5.5"
    p1 = p1 & "-.7.3-1.3.6-1.8 1.2-.6.6-1 1.1-1.2 1.8-.3.7-.5 1.4-.5 2.5-.1 1.1-.1 1.5-.1 4.2s0 3.1.1 4.2c.1 1.1.2 1.8.5 2.5"
    p1 = p1 & ".3.7.6 1.3 1.2 1.8.6.6 1.1 1 1.8 1.2.7.3 1.4.5 2.5.5 1.1.1 1.5.1 4.2.1s3.1 0 4.2-.1c1.1-.1 1.8-.2 2.5-.5"
    p1 = p1 & ".7-.3 1.3-.6 1.8-1.2.6-.6 1-1.1 1.2-1.8.3-.7.5-1.4.5-2.5.1-1.1.1-1.5.1-4.2s0-3.1-.1-4.2c-.1-1.1-.2-1.8-.5-2.5"
    p1 = p1 & "-.3-.7-.6-1.3-1.2-1.8-.6-.6-1.1-1-1.8-1.2-.7-.3-1.4-.5-2.5-.5-1.1-.1-1.5-.1-4.2-.1z"
    
    p2 = "M24 18.9c-2.8 0-5.1 2.3-5.1 5.1s2.3 5.1 5.1 5.1 5.1-2.3 5.1-5.1-2.3-5.1-5.1-5.1zm0 8.4c-1.8 0-3.3-1.5-3.3-3.3"
    p2 = p2 & "s1.5-3.3 3.3-3.3 3.3 1.5 3.3 3.3-1.5 3.3-3.3 3.3z"
    
    svg = "<svg viewBox=""0 0 48 48""><circle fill=""#284b66"" cx=""24"" cy=""24"" r=""24""/>"
    svg = svg & "<path fill=""#fff"" d=""" & p1 & """/>"
    svg = svg & "<path fill=""#fff"" d=""" & p2 & """/>"
    svg = svg & "<circle fill=""#fff"" cx=""29.4"" cy=""18.6"" r=""1.2""/></svg>"
    
    GetInstagramSVG = svg
End Function

Private Function GetFacebookSVG() As String
    Dim p As String
    p = "M26.5 38V25.8h4.1l.6-4.8h-4.7v-3.1c0-1.4.4-2.3 2.4-2.3h2.5V11c-.4-.1-1.9-.2-3.6-.2"
    p = p & "-3.5 0-6 2.1-6 6.1v3.4H17v4.8h4.8V38h4.7z"
    GetFacebookSVG = "<svg viewBox=""0 0 48 48""><circle fill=""#284b66"" cx=""24"" cy=""24"" r=""24""/>" & _
                     "<path fill=""#fff"" d=""" & p & """/></svg>"
End Function

Private Function GetPinterestSVG() As String
    Dim svg As String
    Dim p1 As String, p2 As String, p3 As String, p4 As String, p5 As String
    
    p1 = "M0 500c2.6-141.9 52.7-260.4 150.4-355.4S364.6 1.3 500 0c145.8 2.6 265.3 52.4 358.4 149.4 93.1 97 140.3 "
    p1 = p1 & "213.9 141.6 350.6-2.6 140.6-52.7 258.8-150.4 354.5-97.7 95.6-214.2 144.1-349.6 145.4-46.9 0-93.7-7.2-140.6-21.5 "
    
    p2 = "9.1-14.3 18.2-30.6 27.3-48.8 10.4-22.1 23.4-63.8 39.1-125 3.9-16.9 9.8-39.7 17.6-68.4 9.1 15.6 24.7 29.9 46.9 "
    p2 = p2 & "43 58.6 27.3 120.4 24.7 185.5-7.8 67.7-39.1 114.6-99.6 140.6-181.6 23.4-85.9 20.5-165.7-8.8-239.2C778.3 277 "
    
    p3 = "725.9 224 650.4 191.4c-95-27.3-187.5-24.4-277.3 8.8s-152.3 90.2-187.5 170.9C176.5 401 171 430.7 169 460c-2 "
    p3 = p3 & "29.3-1 57.9 2.9 85.9s13.7 53.1 29.3 75.2 36.5 39.1 62.5 50.8c6.5 2.6 11.7 2.6 15.6 0 5.2-2.6 10.4-13 15.6-31.2 "
    
    p4 = "5.2-18.2 7.2-30.6 5.9-37.1-1.3-2.6-3.9-7.2-7.8-13.7-27.3-44.3-36.5-90.8-27.3-139.6 9.1-48.8 29.3-90.2 60.5-124 "
    p4 = p4 & "48.2-43 104.5-66.4 168.9-70.3 64.4-3.9 119.5 13.7 165 52.7 24.7 28.6 40.7 63.1 47.8 103.5s7.2 79.1 0 116.2c-7.2 "
    
    p5 = "37.1-19.9 71.9-38.1 104.5-32.6 50.8-71 76.8-115.2 78.1-26-1.3-47.2-11.4-63.5-30.3s-21.2-40.7-14.6-65.4c2.6-14.3 "
    p5 = p5 & "10.4-42.3 23.4-84 13-41.7 20.2-72.9 21.5-93.7-3.9-49.5-26.7-74.9-68.4-76.2-32.6 3.9-56.6 18.6-72.3 43.9s-24.1 "
    p5 = p5 & "54.4-25.4 86.9c3.9 37.8 9.8 63.8 17.6 78.1-14.3 58.6-25.4 105.5-33.2 140.6-2.6 9.1-9.8 37.1-21.5 84s-18.2 "
    p5 = p5 & "82.7-19.5 107.4V957C206.3 914 133.3 851.9 80 770.5 26.7 689.1 0 598.9 0 500z"
    
    svg = "<svg viewBox=""0 0 999.9 999.9""><circle cx=""500"" cy=""500"" r=""500"" fill=""#284b66""/>"
    svg = svg & "<path fill=""#fff"" d=""" & p1 & p2 & p3 & p4 & p5 & """/></svg>"
    
    GetPinterestSVG = svg
End Function

Private Function GetHouzzSVG() As String
    GetHouzzSVG = "<svg viewBox=""0 0 48 48""><circle fill=""#284b66"" cx=""24"" cy=""24"" r=""24""/>" & _
                  "<path fill=""#fff"" d=""M30 35h-5v-9h-2v9h-5V22l6-5 6 5v13zm2-14l-8-7-8 7v15h6v-9h4v9h6V21z""/></svg>"
End Function

Private Function GetCalComSVG() As String
    Dim p As String
    p = "M32 16v16H16V16h16m2-2H14v20h20V14zm-4 14h-2v-2h2v2zm0-4h-2v-2h2v2zm0-4h-2v-2h2v2z"
    p = p & "m-4 8h-2v-2h2v2zm0-4h-2v-2h2v2zm0-4h-2v-2h2v2zm-4 8h-2v-2h2v2zm0-4h-2v-2h2v2zm0-4h-2v-2h2v2z"
    GetCalComSVG = "<svg viewBox=""0 0 48 48""><circle fill=""#284b66"" cx=""24"" cy=""24"" r=""24""/>" & _
                   "<path fill=""#fff"" d=""" & p & """/></svg>"
End Function

'-----------------------------------------------------------------------------
' GetAnalyticsScript
' Purpose: Generate Google Analytics tracking code
' Parameters:
'   wsSettings - Settings worksheet
'   cardOwner - Name of card owner
'   primaryRole - Primary job title
'   franchiseCode - Franchise code
' Returns: Google Analytics script tags
'-----------------------------------------------------------------------------
Private Function GetAnalyticsScript(ByVal wsSettings As Worksheet, ByVal cardOwner As String, ByVal primaryRole As String, ByVal franchiseCode As String) As String
    Dim script As String
    Dim propertyID As String
    
    ' Get GA4 property ID from settings
    propertyID = Module5_Utilities.GetSettingValue("Corporate GA4 Property ID", "")
    
    If propertyID = "" Then
        GetAnalyticsScript = ""
        Exit Function
    End If
    
    ' Build GA4 tracking script
    script = "    <!-- Google Analytics -->" & vbCrLf
    script = script & "    <script async src=""https://www.googletagmanager.com/gtag/js?id=" & propertyID & """></script>" & vbCrLf
    script = script & "    <script>" & vbCrLf
    script = script & "        window.dataLayer = window.dataLayer || [];" & vbCrLf
    script = script & "        function gtag(){dataLayer.push(arguments);}" & vbCrLf
    script = script & "        gtag('js', new Date());" & vbCrLf
    script = script & "        gtag('config', '" & propertyID & "', {" & vbCrLf
    script = script & "            'card_owner': '" & Replace(cardOwner, "'", "\'") & "'," & vbCrLf
    script = script & "            'primary_role': '" & Replace(primaryRole, "'", "\'") & "'," & vbCrLf
    script = script & "            'franchise_code': '" & Replace(franchiseCode, "'", "\'") & "'" & vbCrLf
    script = script & "        });" & vbCrLf
    script = script & "    </script>" & vbCrLf
    
    GetAnalyticsScript = script
End Function

'-----------------------------------------------------------------------------
' URLEncode
' Purpose: URL-encode text for data URI
' Parameters:
'   text - Text to encode
' Returns: URL-encoded string
'-----------------------------------------------------------------------------
Private Function URLEncode(ByVal text As String) As String
    Dim i As Long
    Dim char As String
    Dim result As String
    
    result = ""
    
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        
        Select Case char
            Case " "
                result = result & "%20"
            Case vbCr
                result = result & "%0D"
            Case vbLf
                result = result & "%0A"
            Case ":"
                result = result & "%3A"
            Case ";"
                result = result & "%3B"
            Case ","
                result = result & "%2C"
            Case Else
                result = result & char
        End Select
    Next i
    
    URLEncode = result
End Function
