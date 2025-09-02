Option Compare Database
Option Explicit

Public vFrom As String
Public vReplyTo As String
Public vSMTPServer As String
Public vSendUserName As String
Public vSendPW As String
Public vFooter As String
Public vIntro As String


Sub SetGlobalVariables()

    vReplyTo = "Locarno Sailing Club <technology@clublocarno.com>"
    
    vFrom = "Locarno Sailing Club <clublocarnotech@gmail.com>"
    vSMTPServer = "smtp.gmail.com"
    vSendUserName = "clublocarnotech@gmail.com"
'    vSendPW = "Locarno8!"
    vSendPW = "jxdxkbenmcsnsyua"
    
'    vFrom = "technology@clublocarno.com"
'    vSMTPServer = "seafire.mdswireless.com"
'    vSendUserName = "technology@clublocarno.com"
'    vSendPW = "Nacra18!"
    
    vIntro = "Locarno's Door Access system, which provides you access to a number of the doors in the club, is tightly integrated with the Jericho Sailing Centre Association's system.  "
    
    vFooter = "This was an automated message from the Locarno Sailing Club.  Please contact either the Membership Director (Membership@clublocarno.com) or the IT Director (Technology@clublocarno.com) if you have any questions." & vbCrLf & vbCrLf _
        & "The Locarno Sailing Club"
                
End Sub


Sub EmailMembersWithProblems()

    Dim dbThis As Database
    Dim rsMembers As Recordset
    Dim strSQL As String
    Dim EmailStatus As String

    Set dbThis = CurrentDb
    
' Only select members that have problems
' Membership status = Active, JSCA Card is Invalid or missing
    strSQL = "SELECT * FROM tbl_Members WHERE [Membership status]='Active' AND ([JSCACardStatus]='Invalid' OR [JSCACardStatus] Is Null) ORDER BY [First name], [Last name]"
    Debug.Print strSQL
    Set rsMembers = dbThis.OpenRecordset(strSQL)
        
    If rsMembers.RecordCount = 0 Then
        Exit Sub
    End If
        
    Do While Not rsMembers.EOF       ' avoids MoveFirst, as will choke for no records
    
        Debug.Print "Emailing: " & rsMembers![First name] & " " & rsMembers![Last name]
' There is no card on Wild Apricot
        If rsMembers![JSCACardNum] = 0 Or IsNull(rsMembers![JSCACardNum]) = True Then
            With rsMembers
                .Edit
                !ProcessingStatus = "CardError: Card Num is blank in Profile.  "
                .Update
            End With

            EmailStatus = EmailMember(rsMembers![Member ID], "Card Num is blank in Profile")
            With rsMembers
                .Edit
                !EmailSent = EmailStatus
                .Update
            End With

' The card number on Wild Apricot was not 5 digits in length
        ElseIf Len(rsMembers![JSCACardNum]) > 5 Then
            With rsMembers
                .Edit
                !ProcessingStatus = "CardError: Card Num is greater then 5 digits in Profile.  "
                .Update
            End With
            
            EmailStatus = EmailMember(rsMembers![Member ID], "Card Num is greater then 5 digits in Profile")
            With rsMembers
                .Edit
                !EmailSent = EmailStatus
                .Update
            End With
        
' Card was provided but it did not match JSCA records according to criteria
        ElseIf IsNull(rsMembers![JSCACardStatus]) = True Then
            With rsMembers
                .Edit
                !ProcessingStatus = "CardError: Unable to Match with JSCA Records.  "
                .Update
            End With
            
            EmailStatus = EmailMember(rsMembers![Member ID], "Unable to Match with JSCA Records")
            With rsMembers
                .Edit
                !EmailSent = EmailStatus
                .Update
            End With
        
' Cards matchs with JSCA records, but is Invalid
        ElseIf rsMembers![JSCACardStatus] = "Invalid" Then
            With rsMembers
                .Edit
                !ProcessingStatus = "CardError: JSCA marked Card as Invalid.  "
                .Update
            End With
            
            EmailStatus = EmailMember(rsMembers![Member ID], "JSCA marked Card as Invalid")
            With rsMembers
                .Edit
                !EmailSent = EmailStatus
                .Update
            End With
' Process to remove Privileges occurs later in process
            
' Capture all other situations and set Error to True.
' There shouldn't be other situations, but just in case
        Else
            With rsMembers
                .Edit
                !ProcessingStatus = "CardError: Other.  "
                .Update
            End With
        End If
            
        ' get next current member
        rsMembers.MoveNext

    Loop
    
    rsMembers.Close
            
End Sub


Sub test()

    Dim TestEmail As String
' Oliver's WA number: 4229850 - record must be in tbl_Members

'    TestEmail = EmailMember(4229850, "Card Num is blank in Profile")
'    TestEmail = EmailMember(4229850, "Card Num is greater then 5 digits in Profile")
'    TestEmail = EmailMember(12561325, "JSCA marked Card as Invalid")
'    TestEmail = EmailMember(4229850, "Unable to Match with JSCA Records")
'    TestEmail = EmailMember(4229850, "Card has been registered")
'    TestEmail = EmailMember(4229850, "New Card has been registered")
    
    SetGlobalVariables
    SendEmail "Oliver_Thompson@yahoo.ca", "", "Message", "Subject: Test"
    
'    Debug.Print TestEmail
    
End Sub


Function EmailMember(vMemberID As Long, vSubject As String) As String

    Dim objEmail As Object
    Dim objMessage As Object
    
On Error GoTo Err_Process
    
    Dim vDateLastEmailSent As Date
    Dim vMEssage As String
    Dim vMessage2 As String
    Dim vCardNum As String
    Dim vLastName As String
    Dim vFirstName As String
    Dim vEmail As String
    Dim vHelp As String
    Dim strSQL As String
    Dim vDays As String
    
    SetGlobalVariables
    
    vHelp = "Please note that updates to our Door Access System are done nightly and will not be done the day of your changes.  Thus, please plan accordingly.  For more information on how to register your card, visit http://www.clublocarno.com/LocarnoDoorAccess." & vbCrLf & vbCrLf

' Only send same email once per # of vDays
   ' Set number of days delayed.
    vDays = 21
   ' Get last email sent, if any
    vDateLastEmailSent = Nz(DLookup("Max([DateSent])", "tbl_Emails", "[MemberID] = " & vMemberID & " AND [Reason]='" & vSubject & "'"))
    
    vEmail = Nz(DLookup("[e-Mail]", "tbl_Members", "[Member ID]='" & vMemberID & "'"))
    
    If DateDiff("d", vDateLastEmailSent, Now()) > vDays Then
    
        vFirstName = DLookup("[First name]", "tbl_Members", "[Member ID]='" & vMemberID & "'")
        vLastName = DLookup("[Last name]", "tbl_Members", "[Member ID]='" & vMemberID & "'")
            
        Select Case vSubject
            Case "Card Num is blank in Profile"
                vMEssage = "Welcome to the Locarno Sailing Club.  " & vbCrLf & vbCrLf _
                    & "As part of your orientation, you will be advised on how you will use your Jericho Sailing Centre Association Access Card to also access Locarno's facilities.  " _
                    & "Registering your card is easy, is done by updating your membership profile on www.clublocarno.com.  " _
                    & vHelp _
                    & "Note that your card will not be activated until you've completed all the new membership requirements. "

            Case "Card Num is greater then 5 digits in Profile"
                vMEssage = vIntro & vbCrLf & vbCrLf _
                    & "We tried to update your Jericho Sailing Centre Association Access Card in the Locarno Door Access system, but the number is greater than 5 digits long in your membership profile on www.clublocarno.com.  " _
                    & "Please update your information to only include the first 5 digits. " & vbCrLf & vbCrLf _
                    & vHelp
            
            Case "JSCA marked Card as Invalid"
                vMEssage = vIntro & vbCrLf & vbCrLf _
                    & "We tried to update your Jericho Sailing Centre Association (JSCA) Access Card in the Locarno Door Access system, but your card has been marked as 'Invalid' by the JSCA administration.  Please contact them to resolve.  " _
                    & "They can be reached at (604) 224-4177 / admin@jsca.bc.ca." & vbCrLf & vbCrLf _
                    & "Please note that your access to the Locarno facilities will remain inactive until such time as this issue is resolved. " & vbCrLf & vbCrLf _
                    & "Your access card may be invalid for the following reasons.  " & vbCrLf _
                    & "  - You have not yet paid JSCA fees.  To resolve this, you can either pay online (see your member profile at www.clublocarno.com) and send a cheque to JSCA, or by calling/visiting the JSCA office." & vbCrLf _
                    & "  - You have not signed JSCA's waiver form (see #2 on the All Access Checklist at http://www.clublocarno.com/AllAccessChecklist" & vbCrLf & vbCrLf _
                    & vHelp
    
            Case "Unable to Match with JSCA Records"
                vCardNum = DLookup("[JSCACardNum]", "tbl_Members", "[Member ID]='" & vMemberID & "'")
    
                vMEssage = vIntro & vbCrLf & vbCrLf _
                    & "We tried to update your Jericho Sailing Centre Association Access Card in the Locarno Door Access system, but we were unable to match your card number with the Jericho system.  " _
                    & "Please ensure your card number is entered correctly.  " & vbCrLf & vbCrLf _
                    & "The number you provided was: " & vCardNum & ".  (Note: this number should only be the first 5 visible digits on the back of your card.)" & vbCrLf & vbCrLf _
                    & "If this is correct, note that your first initial and full last name need to match exactly with what JSCA has on file.  " _
                    & "Please either update your records on the Locarno Membership system (www.clublocarno.com), or contact the JSCA office at (604) 224-4177 / admin@jsca.bc.ca to confirm the following information they have on file exactly matches what you provided to us: " & vbCrLf _
                    & "     Last Name: " & vLastName & vbCrLf _
                    & "     First Initial: " & Left(vFirstName, 1) & vbCrLf _
                    & "     Club affiliation: Locarno Sailing Club" & vbCrLf & vbCrLf _
                    & "Please note that your access to the Locarno facilities will remain inactive until such time as this issue is resolved. " & vbCrLf & vbCrLf _
                    & vHelp
    '            Debug.Print vMessage
    
            Case "Card has been registered"
                vMEssage = vIntro & vbCrLf & vbCrLf _
                    & "Your Jericho Sailing Centre Association Access Card has been registered in the Locarno Door Access system, and you will have access to the fleet(s) you have signed up for once you had completed the new member orientation." & vbCrLf
            
            Case "New Card has been registered"
                vMEssage = "Your new Jericho Sailing Centre Association Access Card has been registered in the Locarno Door Access system, and should be available to use." & vbCrLf
        
        End Select
        
        Debug.Print "Emailing: " & vFirstName & " " & vLastName
        
        vMessage2 = vFirstName & ":" & vbCrLf & vbCrLf & vMEssage
        
        SendEmail vEmail, "", vMessage2, "Message from the Locarno Sailing Club: " & vSubject
'        SendEmail vEmail, vReplyTo, vMessage2, "Message from the Locarno Sailing Club: " & vSubject
        Debug.Print "Email Sent Successfully."
        
        ' DoCmd.SetWarnings False
        
        strSQL = "INSERT INTO tbl_Emails ( MemberID, ToName, ToEmail, Reason, DateSent) " _
            & "VALUES (" & vMemberID & ", '" & vFirstName & " " & vLastName & "','" & vEmail & "','" & vSubject & "','" & Now() & "')"
        Debug.Print strSQL
        DoCmd.RunSQL (strSQL)
        
        EmailMember = "Email was sent."
    
    Else
    
        EmailMember = "Email was NOT sent as one was sent within the last " & vDays & " days."
    
    End If
    
        
Exit_Process:
    Exit Function
    
Err_Process:

    EmailMember = "Error: Email was NOT sent."
    Resume Exit_Process
    
End Function


Sub EmailStatus()

On Error GoTo Err_Process

    Dim dbThis As Database
    Dim rsMembers As Recordset
    Dim rsSystems As Recordset
    Dim vMEssage As String
    Dim vMessage2 As String
    Dim LastUpdate As Date
    Dim strSQL As String
    Dim ErrorDescr As String
    Dim vTo As String
    Dim Alert As Boolean
    Dim AlertMsg As String
    Dim vPreviousStatus As String
    Dim vDoorAccessDisableReason As String
    Dim x As Integer
    Dim NumPrivileges As Integer

    Set dbThis = CurrentDb
    
    strSQL = "SELECT * FROM tbl_Members WHERE [Membership Status] <> 'Lapsed' ORDER BY [Membership Status], ProcessingStatus, [First name]"
    Set rsMembers = dbThis.OpenRecordset(strSQL)
    
    If rsMembers.RecordCount = 0 Then
        Exit Sub
    End If
        
    SetGlobalVariables
    
    LastUpdate = Format(DLookup("Max(LastUpdate)", "tbl_LastUpdate"), "yyyy-mm-dd hh:mm")
 
    vMEssage = "The Locarno Club Access System has been updated with records from Wild Apricot as of: " & LastUpdate & "." & vbCrLf & vbCrLf _
        & "The following members' records have been processed."
        
    vPreviousStatus = ""
    x = 1

' Walk through current members
    Do While Not rsMembers.EOF       ' avoids MoveFirst, as will choke for no records
    
        If rsMembers![Membership Status] & rsMembers![ProcessingStatus] <> vPreviousStatus Then
            vMEssage = vMEssage & vbCrLf & vbCrLf & rsMembers![Membership Status] & " Members: " & rsMembers![ProcessingStatus] & vbCrLf & vbCrLf
            x = 1
        End If
        
        vMEssage = vMEssage & "  " & x & ".  " & rsMembers![First name] & " " & rsMembers![Last name] & "    " & rsMembers![EmailSent]
        
        If rsMembers![Door Access] = "Disabled" Then
            If IsNull(rsMembers![DoorAccessDisableReason]) Or rsMembers![DoorAccessDisableReason] = "" Then
                vDoorAccessDisableReason = "None Provided on WA"
            Else
                vDoorAccessDisableReason = rsMembers![DoorAccessDisableReason]
            End If
            vMEssage = vMEssage & "   (Door Disabled. Reason: " & vDoorAccessDisableReason & ")"
        End If
            
        vMEssage = vMEssage & vbCrLf

        vPreviousStatus = rsMembers![Membership Status] & rsMembers![ProcessingStatus]
        rsMembers.MoveNext
        x = x + 1
    
    Loop
    
    vMEssage = vMEssage & vbCrLf & vbCrLf & "==============================" & vbCrLf & vbCrLf


' Confirm status of connectivity
    
    Alert = False
    
    strSQL = "SELECT tbl_Connectivity.Description, tbl_Connectivity.IPAddress FROM tbl_Connectivity ORDER BY tbl_Connectivity.SortOrder"
'    Debug.Print strSQL
    Set rsSystems = dbThis.OpenRecordset(strSQL)  ' only contains MemberID and FullName
        
    vMessage2 = "The following is the status of Locarno's systems:" & vbCrLf
    
    ' Walk through devices
    Do While Not rsSystems.EOF       ' avoids MoveFirst, as will choke for no records
    
        vMessage2 = vMessage2 & "   " & rsSystems![Description] & ": "
        If SystemOnline(rsSystems![IPAddress]) = True Then
            vMessage2 = vMessage2 & ": On" & vbCrLf
        Else
            vMessage2 = vMessage2 & ": Connectivty failed" & vbCrLf
            Alert = True
        End If
    
        rsSystems.MoveNext
   
    Loop
    
    rsSystems.Close
    
    vMEssage = vMEssage & vbCrLf & vMessage2 & vbCrLf
    
    
' Get number of members that have priviledges
'    strSQL = "SELECT Count(*) AS N FROM (SELECT DISTINCT  t_d_Privilege.f_ConsumerID FROM t_d_Privilege) AS T;"
    
    NumPrivileges = DCount("f_ConsumerID", "t_d_Privilege")
   
    vMEssage = vMEssage & "There are " & NumPrivileges & " door access privileges in the system."
    
    Debug.Print vMEssage

    vTo = "Technology@clublocarno.com, Membership@clublocarno.com"
    SendEmail vTo, "", vMEssage, "Locarno Card System Status Update: " & LastUpdate & AlertMsg
    Debug.Print "Update Email Sent Successfully."


Exit_Process:
    Exit Sub

Err_Process:
    Debug.Print Err.Description
    ErrorDescr = Replace(Err.Description, "'", " ")
    strSQL = "INSERT INTO tbl_ErrorLog (ProcessStep, Description) VALUES ('Sub EmailStatus', '" & ErrorDescr & "')"
    DoCmd.RunSQL (strSQL)
        Resume Exit_Process

End Sub


Sub SendFailedDownloadMessage(vErrorMessage As String)

Dim vTo As String
Dim vSubject As String

    SetGlobalVariables
    
    vTo = "Technology@clublocarno.com"
    vSubject = "Download failed to complete."
    
    SendEmail vTo, "", vErrorMessage, vSubject

End Sub


Sub SendEmail(vTo As String, vBCC As String, vMEssage As String, vSubject As String)
    
Dim objEmail As Object
Dim schema As String
    
    Set objEmail = CreateObject("CDO.Message")  'The space between Create and Object is on purpose. Otherwise the text editor of this forum gives an error.
    
    schema = "http://schemas.microsoft.com/cdo/configuration/"
    
    With objEmail
        .from = vFrom
        .To = vTo
'        .Bcc = vBCC
        .ReplyTo = vReplyTo
        .Subject = vSubject
        .Textbody = vMEssage & vbCrLf & vbCrLf & vFooter
     
        With .Configuration.Fields
            .Item(schema & "sendusing") = 2
            .Item(schema & "smtpserver") = vSMTPServer
            .Item(schema & "smtpserverport") = 465
            .Item(schema & "smtpauthenticate") = 1
            .Item(schema & "sendusername") = vSendUserName
            .Item(schema & "sendpassword") = vSendPW
            .Item(schema & "smtpconnectiontimeout") = 30
            .Item(schema & "smtpusessl") = 1
        End With
    
        .Configuration.Fields.Update
        .send
    End With
    
'    Debug.Print "Email send function is OFF"


End Sub




Function SystemOnline(ByVal ComputerName As String)
' Old code - not used.  Replaced with Network Monitoring tool

' This function returns True if the specified IP address can be pinged.
' The Win32_PingStatus class used in this function requires Windows XP or later.
' Standard housekeeping
    Dim colPingResults As Variant
    Dim oPingResult As Variant
    Dim strQuery As String

' Define the WMI query
    strQuery = "SELECT * FROM Win32_PingStatus WHERE Address = '" & ComputerName & "'"

' Run the WMI query
    Set colPingResults = GetObject("winmgmts://./root/cimv2").ExecQuery(strQuery)
' Translate the query results to either True or False
    For Each oPingResult In colPingResults
        If Not IsObject(oPingResult) Then
            SystemOnline = False
        ElseIf oPingResult.StatusCode = 0 Then
            SystemOnline = True
        Else
            SystemOnline = False
        End If
    Next

End Function