Option Compare Database

Public OAuthToken As String
Public OAuthUrl As String
Public ApiKey As String
Public ApiUrl As String
Public ContactsResultUrl As String
Public ContactFields As IXMLDOMSelection

' DELETE Public ExcludedFields(200) As String
' DELETE Public NumberOfColumns As Integer

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Sub GetDatafromWP(vScope As String)
 On Error GoTo ErrorHandler
    
    ContactsResultUrl = ""
    ApiKey = "APIkeyvalue"
    ApiUrl = "https://api.wildapricot.org"
    OAuthUrl = "https://oauth.wildapricot.org/auth/token"
    
    Debug.Print ("authorization")
    OAuthToken = GetOAuthToken(OAuthUrl)
    
    Debug.Print ("start downloading process")
    
    Set xmlDoc = LoadXml(ApiUrl)
    Set apiVersion = xmlDoc.SelectSingleNode("//ApiVersion/Version")
    Set versionUrl = xmlDoc.SelectSingleNode("//ApiVersion/Url")
    accountUrl = LoadAccountUrl(versionUrl.text)
    Set accountInfoDoc = LoadXml(accountUrl)
   
' Toast    ApiUrl = "https://api.wildapricot.org/v1/accounts/38348/Contacts/?"

    contactsBaseUrl = accountInfoDoc.SelectSingleNode("//Resources/Resource[Name='Contacts']/Url").text
    LoadContacts contactsBaseUrl, vScope
    
    Exit Sub
    
ErrorHandler:
    MsgBox ("Something goes wrong:" + Err.Description)
End Sub


Sub LoadContacts(ByVal baseUrl As String, vScope As String)
    Dim url As String
    Dim filter As String, selectExpression As String
    Dim LastUpdateTime As String
   
    selectExpression = "$select='Member ID', 'JSCA Card Number','First name','Last name','e-Mail','Membership Status'," _
        & "'Membership level ID','Renewal due'," _
        & "'All Fleets','Sailing Fleet','Windsurfing Fleet','Kayaking Fleet','Rowing Fleet','SUP Stand Up Paddle Fleet'," _
        & ",'Windsurfing Level 3','Door Access', 'Door Access Disable Reason'"

'        & ",'Access to Level 3 Storage'"
'        & "'custom-8152854','custom-7532343','custom-7532345','custom-7532340','custom-7532341','custom-7532344'," _

    url = baseUrl
        
    If vScope = "Incremental" Then ' add filter
        LastUpdate = Format(DLookup("Max(LastUpdate)", "tbl_LastUpdate"), "yyyy-mm-dd")
        LastUpdateTime = Format(DLookup("Max(LastUpdate)", "tbl_LastUpdate"), "hh:mm:ss")
        
        LastUpdate = LastUpdate & "T" & LastUpdateTime
        filter = "$filter='Profile last updated' ge '" & LastUpdate & "'"
    
    Else ' full load
        filter = "$filter='Membership status' eq 'Active' AND 'Process Door Access?' ne 'No' "
    '    filter = "$filter='MemberID eq 33020847' &"
    
    End If
    
    If Len(filter) > 0 Then
        url = url + "?" + filter
    End If
    
    If Len(selectExpression) > 0 Then
        If Len(filter) > 0 Then
            url = url + "&" + selectExpression
        Else
            url = url + "?" + selectExpression
        End If
    End If
    
    Debug.Print url;
    
    Dim contactsRequestDoc As DOMDocument
    Set contactsRequestDoc = LoadXml(url)
    ContactsResultUrl = contactsRequestDoc.SelectSingleNode("/ApiResponse/ResultUrl").text
    CheckContactsResult

End Sub


Sub CheckContactsResult()
    Dim complete As Boolean, requestState As String
    complete = False
    Dim contactsRequestDoc As DOMDocument
    Dim x As Integer
    
    Set contactsRequestDoc = LoadXml(ContactsResultUrl)
    requestState = contactsRequestDoc.SelectSingleNode("/ApiResponse/State").text
    Debug.Print requestState
    
    x = 0
    
    If requestState = "Waiting" Or requestState = "Processing" Then
        x = x + 1
        Sleep (10000)   ' Waits for 10 seconds,
        Debug.Print "Waiting for Response"
        CheckContactsResult
        If x = 10 Then
            SendFailedDownloadMessage ("Failed to load contacts. Timeout.")
            complete = True
            Exit Sub
        End If
    
    ElseIf requestState = "Complete" Then
        Dim contacts As IXMLDOMSelection
        Set contacts = contactsRequestDoc.SelectNodes("/ApiResponse/Contacts/Contact")
        
        strSQL = "DELETE * FROM tbl_Members"
        DoCmd.RunSQL (strSQL)
      
        SaveContacts contacts
        complete = True
    
        strSQL = "Insert Into tbl_LastUpdate (LastUpdate) Values ('" & Now() & "')"
        DoCmd.RunSQL (strSQL)
    
    ElseIf requestState = "Failed" Then
        Dim errMessage As String
        errMessage = contactsRequestDoc.SelectSingleNode("/ApiResponse/ErrorDetails").text
'        MsgBox ("Failed to load contacts. Error: " + errMessage)
        SendFailedDownloadMessage ("Failed to load contacts. Error: " & errMessage)
        complete = True
    
    End If
End Sub


Sub test()
        SendFailedDownloadMessage ("Failed to load contacts. Error: ")

End Sub

Sub SaveContacts(contacts As IXMLDOMSelection)

On Error GoTo Err_Process

    Dim contact As IXMLDOMElement
    Dim CountTotalAPI As Integer
    Dim CountTotalSaved As Integer
    Dim vMembershipStatus As String
    Dim vEmail As String
    Dim vMembershipLevel As String
    Dim vFleets As String
    Dim vWSL3 As String
    Dim vRenewalDue As String
    Dim vDoorAccess As String
    Dim vDoorAccessDisableReason As String
    
    CountTotalAPI = 0
    CountTotalSaved = 0
    For Each contact In contacts
        CountTotalAPI = CountTotalAPI + 1
        
        Dim valueNode As IXMLDOMElement
        Dim fieldNode As IXMLDOMElement
        
        Set valueNode = Nothing
        Set fieldNodes = contact.SelectNodes("FieldValues/ContactField")
        
        For Each fieldNode In fieldNodes
            Dim localFieldName As IXMLDOMElement
            Set localFieldName = fieldNode.SelectSingleNode("FieldName")
                      
            Debug.Print fieldNode.SelectSingleNode("FieldName").text ' & ",: " & fieldNode.SelectSingleNode("FieldName").Value
           
' Insert into tbl_Members ([Member ID], [JSCACardNum], [JSCACardNum_Text], [First name], [Last name], [e-Mail],               [Membership Status],      [Membership Level],      [Renewal due],   [Fleets],  [Windsurfing Level 3], [Door Access],      [DoorAccessDisableReason])
' V1 Values               (33020847,     52529,             52529,          'Ryan',        'Alban',   ' thegrogen@gmail.com',   'Active',               'Regular Member/Other',  '2018-06-17',    'Sailing', '',                     'Enabled',          '')
' V2 Values               (33020847,     52529,             52529,          'Ryan',        'Alban',   ' thegrogen@gmail.com',   'Active',               'Regular Member/Other',  '2018-06-18',    'Sailing',  '',                    'Enabled',          '')

             
            Select Case fieldNode.SelectSingleNode("FieldName").text
                Case "Member ID"
                    vMemberID = fieldNode.SelectSingleNode("Value").text
                Case "JSCA Card Number"
                    vJSCACardNum = fieldNode.SelectSingleNode("Value").text
                Case "First name"
                    vFirstName = fieldNode.SelectSingleNode("Value").text
                Case "Last name"
                    vLastName = fieldNode.SelectSingleNode("Value").text
'                    Debug.Print "Last Name: " & fieldNode.SelectSingleNode("Value").Text
                Case "e-Mail"
                    vEmail = fieldNode.SelectSingleNode("Value").text
                Case "Membership Status"
                    If Left(fieldNode.SelectSingleNode("Value").text, 1) = 1 Then
                            vMembershipStatus = "Active"
                    End If
'        "Id": 2,        "Label": "Lapsed",
'        "Id": 20,       "Label": "Pending - New",
'        "Id": 3,        "Label": "Pending - Renewal",
'        "Id": 30,       "Label": "Pending - Level change",
                    
                Case "Membership level ID"
                
'                    Debug.Print "Membership Level ID: " & fieldNode.SelectSingleNode("Value").Text
                    
                    Select Case fieldNode.SelectSingleNode("Value").text
                        Case 162949
                            vMembershipLevel = "Executive"
                        Case 162957
                            vMembershipLevel = "Sponsored"
                        Case 700166
                            vMembershipLevel = "Instructor"
                        Case Else
                            vMembershipLevel = "Regular Member/Other"
'                           Case 162952 = "Regular Member"
'                           Case 162951 = "Fleet Captain"
                        End Select
                Case "Renewal due"
                    vRenewalDue = Left(fieldNode.SelectSingleNode("Value").text, 10)
                
                Case "Fleets"
                    vFleets = fieldNode.SelectSingleNode("Value").text
                
                Case "All Fleets"
                    If Right(fieldNode.SelectSingleNode("Value").text, 3) = "Yes" Then
                        vFleets = vFleets & "All Fleets"
                    End If
                Case "Sailing Fleet"
                    If Right(fieldNode.SelectSingleNode("Value").text, 3) = "Yes" Then
                        vFleets = vFleets & "Sailing"
                    End If
                Case "Windsurfing Fleet"
                    If Right(fieldNode.SelectSingleNode("Value").text, 3) = "Yes" Then
                        vFleets = vFleets & "Windsurfing"
                    End If
                Case "Kayaking Fleet"
                    If Right(fieldNode.SelectSingleNode("Value").text, 3) = "Yes" Then
                    vFleets = vFleets & "Kayaking"
                    End If
                Case "Rowing Fleet"
                    If Right(fieldNode.SelectSingleNode("Value").text, 3) = "Yes" Then
                        vFleets = vFleets & "Rowing"
                    End If
                Case "SUP Stand Up Paddle Fleet"
                    If Right(fieldNode.SelectSingleNode("Value").text, 3) = "Yes" Then
                        vFleets = vFleets & "SUP"
                    End If
             
                Case "Windsurfing Level 3"
                    vWSL3 = fieldNode.SelectSingleNode("Value").text
                Case "Door Access"
                    If Left(fieldNode.SelectSingleNode("Value").text, 7) = "3467458" Then
                        vDoorAccess = "Enabled"
                    Else
                        vDoorAccess = "Disabled"
                    End If
                        
                Case "Door Access Disable Reason"
                    vDoorAccessDisableReason = fieldNode.SelectSingleNode("Value").text
            End Select
        
        Next fieldNode
        
        If vJSCACardNum = "" Then vJSCACardNum = 0
        Debug.Print "Member : " & vFirstName & " " & vLastName & ",  MemberID: " & vMemberID & ",  Membership status: " & vMembershipStatus & ",  Membershipship Level: " & vMembershipLevel
        

' ============
        
' Address WA API bug where not all Member IDs are returned in result.
' If there is no memberID but a status, then look in local table for memberID

        If vMemberID = "" And vMembershipStatus <> "" Then
            vMemberID = Nz(DLookup("[MemberID]", "tbl_LocalMemberIDs", "[FullName]='" & vFirstName & " " & vLastName & "'"))
' If not found, insert name.  Periodically reference table to add number.
            If vMemberID = "" Then
                sqlstr = "Insert into tbl_LocalMemberIDs ([FullName], DateReferenced) Values ('" & vFirstName & " " & vLastName & "', Now())"
                ' DoCmd.SetWarnings False
    '            Debug.Print sqlstr
                DoCmd.RunSQL (sqlstr)
            Else
' If found, add date to show record is still being used
                sqlstr = "UPDATE tbl_LocalMemberIDs SET [DateReferenced] = Now() WHERE [FullName] = '" & vFirstName & " " & vLastName & "'"
                ' DoCmd.SetWarnings False
    '            Debug.Print sqlstr
                DoCmd.RunSQL (sqlstr)
            End If
        End If
        
' ============
        
' API result must have MemberID and Membership Status to save to database
        If IsNumeric(vJSCACardNum) = False Then vJSCACardNum = 0
        If vMemberID <> "" And vMembershipStatus <> "" Then
            CountTotalSaved = CountTotalSaved + 1
            sqlstr = "Insert into tbl_Members " _
                & "([Member ID], [JSCACardNum], [JSCACardNum_Text], [First name], [Last name], [e-Mail], " _
                & "[Membership Status], [Membership Level], [Renewal due], [Fleets], " _
                & "[Windsurfing Level 3], [Door Access], [DoorAccessDisableReason]) " _
                & "Values (" _
                & vMemberID & ", " & vJSCACardNum & ", " & vJSCACardNum & ", '" _
                & vFirstName & "', '" & vLastName & "',' " & vEmail & "','" _
                & vMembershipStatus & "','" & vMembershipLevel & "','" & vRenewalDue & "','" & vFleets & "','" _
                & vWSL3 & "','" & vDoorAccess & "','" & vDoorAccessDisableReason & "')"
            ' DoCmd.SetWarnings False
            Debug.Print sqlstr
            DoCmd.RunSQL (sqlstr)
            
        End If
        
        vMemberID = ""
        vJSCACardNum = ""
        vFirstName = ""
        vLastName = ""
        vEmail = ""
        vMembershipStatus = ""
        vMembershipLevel = ""
        vRenewalDue = ""
        vFleets = ""
        vWSL3 = ""
        vDoorAccess = ""
        
    Next contact
    Debug.Print "Total Records in API: " & CountTotalAPI
    Debug.Print "Total Records Saved: " & CountTotalSaved
    UpdateStatusBar ("Total Records in API: " & CountTotalAPI)
    UpdateStatusBar ("Total Records Saved: " & CountTotalSaved)


Exit_Process:
    Exit Sub

Err_Process:
    Debug.Print Err.Description
    ErrorDescr = Replace(Err.Description, "'", " ")
    strSQL = "INSERT INTO tbl_ErrorLog (ProcessStep, Description) VALUES ('Sub SaveContacts', '" & ErrorDescr & "')"
    DoCmd.RunSQL (strSQL)
        Resume Exit_Process
        
End Sub


' Other code from Wild Apricot



Function EncodeBase64(text As String) As String
  Dim arrData() As Byte
  arrData = StrConv(text, vbFromUnicode)

  Dim objXML As MSXML2.DOMDocument
  Dim objNode As MSXML2.IXMLDOMElement

  Set objXML = New MSXML2.DOMDocument
  Set objNode = objXML.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  EncodeBase64 = objNode.text

  Set objNode = Nothing
  Set objXML = Nothing
End Function

Sub SetOAuthCredentials(httpClient As IXMLHTTPRequest)
    httpClient.setRequestHeader "User-Agent", "VBA sample app" ' This header is optional, it tells what application is working with API
    httpClient.setRequestHeader "Authorization", "Basic " + EncodeBase64("APIKEY:" + ApiKey) ' This header is required, it provides API key for authentication
    httpClient.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
End Sub

Function GetOAuthToken(ByVal url As String) As String
    Debug.Print ("Loading data from " + url)
    Dim httpClient As IXMLHTTPRequest
    Set httpClient = CreateObject("Msxml2.XMLHTTP.3.0")
    httpClient.Open "POST", url, False
    SetOAuthCredentials httpClient
    
    httpClient.send ("grant_type=client_credentials&scope=auto")
    
    If Not httpClient.status = 200 Then
        msg = "Call to " + url + " returned error:" + httpClient.statusText
        Err.Raise -1, "GetOAUthToken", msg
    End If
    
    Dim resp As String
    resp = httpClient.responseText
    resp = Mid(resp, Len("{""access_token"":""") + 1, InStr(resp, """,""token_type""") - Len("{""access_token"":""") - 1)

    GetOAuthToken = resp
End Function

Sub SetCredentials(httpClient As IXMLHTTPRequest)
    httpClient.setRequestHeader "User-Agent", "VBA sample app" ' This header is optional, it tells what application is working with API
    httpClient.setRequestHeader "Authorization", "Bearer " + OAuthToken  ' This header is required, it provides API key for authentication
    httpClient.setRequestHeader "Accept", "application/xml" ' This header is required, it tells to return data in XML format
End Sub

Function LoadXml(ByVal url As String) As DOMDocument
    Debug.Print ("Loading data from " + url)
    Dim httpClient As IXMLHTTPRequest
    Set httpClient = CreateObject("Msxml2.XMLHTTP.3.0")
    httpClient.Open "GET", url, False
    SetCredentials httpClient
    httpClient.send
    
    If Not httpClient.status = 200 Then
        msg = "Call to " + url + " returned error:" + httpClient.statusText
        Err.Raise -1, "LoadXML", msg
    End If

    Set xmlDoc = httpClient.responseXML
    Set LoadXml = xmlDoc
End Function

Function LoadAccountUrl(versionUrl As String) As String
    Set versionResourcesXml = LoadXml(versionUrl)
    Set accountsUrlNode = versionResourcesXml.SelectSingleNode("//Resource[Name='Accounts']/Url")
    LoadAccountUrl = accountsUrlNode.text
End Function