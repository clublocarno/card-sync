Option Compare Database
Option Explicit

Public Today As Date
Public ThisYear As Date
Public SeasonEnd As Date



Public Sub UpdateCardSystem()

' Adds prileges for members:
'    Locarno Membership Status = Active AND
'    Jericho Card = Valid AND
'    Door Access = Enabled
    UpdateCardSystemAdd

' Removes Privleges for Lapsed Members
    RemovePrivilegesLapsedMembers

' Removes Privleges for
'    Locarno Membership Status = Active AND
'    Door Access = NOT Enabled
    RemovePrivilegesDoorAccessNOTEnabled

' Do nothing with Pending Status

' Instructional Emails sent to the following:
'    Locarno Membership status Active AND
'    Jericho Card = Invalid OR is Null


End Sub


Public Sub UpdateCardSystemAdd()
' Syncs card database from membership database,
' including checks with JSCA member/card database.
' Installs Member name, MemberID, start date, expiry date in carddb Personnel.
' Set door privildges based on fleet membership (and in the case of Windsurfing cage,
' windsurfing Level 3 clearance Or specific Level 3 cage clearance
' as reported in the membership db.
' Door privileges auto-expire on each memerbs Renewal Date given in the membership db.
' Door privileges can be removed by setting field "Door Access" = Disable.
' Simly revoke if JSCA card is invalid after a grace period of 15 days from DateCurPaid REVIEW
'
' Created: Version 2.0 March 2011, Alan Kot
' Based on Version 1.0 from 2010, which used the membership
' database using MS Access, that had been in use since 2002


On Error GoTo Err_Process

    Dim strSQL As String
    Dim dbThis As Database
    Dim rsMembers As Recordset
    Dim rsThisMember As Recordset
    Dim rsMembersLapsed As Recordset
    Dim rs As Recordset
    Dim MemberID As Long
    Dim FullName As String
    Dim memberLev As String
    Dim passDes As String
    Dim privMain As Boolean
    Dim privSail As Boolean
    Dim privWind As Boolean
    Dim privWindCage As Boolean
    Dim privKayakRow As Boolean
    Dim DoorID As Integer
    Dim denyPrivs As Boolean
    Dim cardValid  As Boolean
    Dim CardNO As String
    Dim jStatus As String
    Dim ConsumerID As Long
    Dim vRenewalDue As String
    Dim expiryDate As Date
        
    Set dbThis = CurrentDb

    
' Get todays date. Used for example to set expiry date from refunds or deemed no access
    Today = Date
    ThisYear = Year(Today)
     
' Set season dates. No need to set start date.
' 2011 Website provides expiry date in field Renewal Date
' Force Season end to avoid multi-year renewal quirk on website (Sponsored or exec could easily do this.)
    SeasonEnd = DateAdd("yyyy", ThisYear - 2000 + 1, "4/1/2000") ' Odd method but Access dates seem finicky to me. This works fine.

    
    ' Get and walk through current members.  Ignore [Status] = Pending - xxx
    
    strSQL = "SELECT DISTINCT [First name] & ' ' & [Last name] AS FullName, [Member ID] AS MemberID, [Renewal due] " _
        & "FROM tbl_Members WHERE " _
        & "[Membership status] = 'Active' " _
        & "AND [JSCACardStatus] = 'Valid' " _
        & "AND CardRegistered = True " _
        & "AND [Door Access] = 'Enabled' " _
        & "ORDER BY [First name] & ' ' & [Last name] "
    Debug.Print strSQL
    Set rsMembers = dbThis.OpenRecordset(strSQL)  ' only contains MemberID and FullName
        
    If rsMembers.RecordCount = 0 Then
        Exit Sub
    End If
        
    ' Walk through current members
    Do While Not rsMembers.EOF       ' avoids MoveFirst, as will choke for no records
        
        ' get name and MemberID
        FullName = rsMembers![FullName]
        MemberID = rsMembers![MemberID]
        Debug.Print "PROCESSING: FullName = " & FullName & " "
                
        If IsNull(rsMembers![Renewal due]) = True Then
            vRenewalDue = ""
        Else
            vRenewalDue = rsMembers![Renewal due]
        End If

    ' Update/set expiry date
        If IsDate(vRenewalDue) Then
            expiryDate = vRenewalDue ' expiry Date is set by Wild Apricot
        '                If DateDiff("d", expiryDate, SeasonEnd) < 0 Then    ' limit expiryDate to SeasonEnd
        '                    expiryDate = SeasonEnd
        '                End If
        Else ' if Renewal due date is not set, set to end of Season.
            expiryDate = SeasonEnd
        End If
    
        strSQL = "UPDATE t_b_Consumer SET t_b_Consumer.f_EndYMD = '" & expiryDate & "' WHERE t_b_Consumer.f_WorkNo = '" & MemberID & "'"
'    Debug.Print strSQL
        DoCmd.RunSQL (strSQL)



        ' Get this members record, only needed fields TODO
        ' Note that MemberID is a Long, whereas the linked tbl_Members table is a string with excess spaces up front
        strSQL = " SELECT * FROM tbl_Members WHERE ( CLng( tbl_Members.[Member ID]) =" & MemberID & " ); "
        Set rsThisMember = dbThis.OpenRecordset(strSQL)       ' will have one record

        ' If Status is NOT Pending, processs this member. Will process Active
        ' If we sync in the period Jan 1- Mar 31, would be proper to use the members reported Renewal date
        ' We would monkey with their fleets, but won't upload, but useful to register cards.
        If InStr(rsThisMember![Membership Status], "Pending") = 0 Then
            
            ' Door privileges
'toast            denyPrivs = False       ' set this later if need to deny all privs
            'Init all door privs to false, except Main to true since all current members get this
            privMain = True
            privSail = False
            privWind = False
            privWindCage = False
            privKayakRow = False
            
            ' Set door privileges based on fields Fleets and Membership Level
            memberLev = rsThisMember![Membership Level]
            passDes = rsThisMember![fleets]
            Select Case memberLev  ' Switch by memberLev
                Case "Executive"  '  all entry doors, plus the ws cage
                    privSail = True
                    privKayakRow = True
                    privWind = True
                    privWindCage = True
                Case "Sponsored"   '  all entry doors, ws cage if Level3
                    privSail = True
                    privKayakRow = True
                    privWind = True
                    privWindCage = checkWindLevel3(MemberID, rsThisMember)
                Case "Instructor"   '  all entry doors, ws cage if Level3
                    privSail = True
                    privKayakRow = True
                    privWind = True
                    privWindCage = checkWindLevel3(MemberID, rsThisMember)
                Case Else   ' All other levels, e.g. "Regular Membership", "Fleet Captain"
                    ' Inspect [Fleets] strings
                    If InStr(passDes, "Sailing") > 0 Then
                        privSail = True
                    End If
                    
                    If (InStr(passDes, "Rowing") > 0) Or (InStr(passDes, "Kayak") > 0) Or (InStr(passDes, "SUP") > 0) Then
                        privKayakRow = True
                    End If

                    If InStr(passDes, "Windsurfing") > 0 Then
                        privWind = True
                        ' also check Level 3 clearance
                        privWindCage = checkWindLevel3(MemberID, rsThisMember)
                    End If
                    
                    If InStr(passDes, "All fleets") > 0 Then
                        privSail = True
                        privWind = True
                        privKayakRow = True
                        privWindCage = checkWindLevel3(MemberID, rsThisMember)
                    End If
                    
' Added 2014 - All Fleets Category was blank in WA where previously it was 'All Fleets'
                    If passDes = "" Then
                        privSail = True
                        privWind = True
                        privKayakRow = True
                        privWindCage = checkWindLevel3(MemberID, rsThisMember)
                    End If
            End Select
            
            ' Set/delete privilege records according to desired door clearances
            setPriv MemberID, "Main", privMain
            setPriv MemberID, "Sailing", privSail
            setPriv MemberID, "Windsurf Main", privWind
            setPriv MemberID, "Windsurf Cage", privWindCage
            setPriv MemberID, "KayakRow", privKayakRow
            
        End If ' end of If [Status] NOT Pending

        With rsThisMember
            .Edit
            !ProcessingStatus = "Privileges Updated.  "
            .Update
        End With
        
        ' get next current member
        rsMembers.MoveNext

    Loop
    
    rsThisMember.Close
    rsMembers.Close


Exit_Process:
    Exit Sub

Err_Process:
    MsgBox Err.Description
        Resume Exit_Process
    
End Sub


Sub RemovePrivilegesLapsedMembers()

    Dim strSQL As String
    
    ' DoCmd.SetWarnings False
        
    strSQL = "DELETE * FROM t_d_Privilege " _
        & "WHERE  t_d_Privilege.f_RecID IN ( " _
        & "SELECT t_d_Privilege.f_RecID " _
        & "FROM tbl_Members RIGHT JOIN (t_b_Consumer INNER JOIN t_d_Privilege ON t_b_Consumer.f_ConsumerID = t_d_Privilege.f_ConsumerID) ON tbl_Members.[Member ID] = t_b_Consumer.f_WorkNo " _
        & "WHERE [Membership status] = 'Lapsed')"
    Debug.Print strSQL
    DoCmd.RunSQL (strSQL)

    strSQL = "UPDATE tbl_Members " _
        & "SET [ProcessingStatus] = 'Priviledges Removed.  '" _
        & "WHERE [Membership status] = 'Lapsed'"
    Debug.Print strSQL
    DoCmd.RunSQL (strSQL)

End Sub


Sub RemovePrivilegesDoorAccessNOTEnabled()

    Dim strSQL As String
    
    ' DoCmd.SetWarnings False
        
    strSQL = "DELETE * FROM t_d_Privilege " _
        & "WHERE  t_d_Privilege.f_RecID IN ( " _
        & "SELECT t_d_Privilege.f_RecID " _
        & "FROM tbl_Members RIGHT JOIN (t_b_Consumer INNER JOIN t_d_Privilege ON t_b_Consumer.f_ConsumerID = t_d_Privilege.f_ConsumerID) ON tbl_Members.[Member ID] = t_b_Consumer.f_WorkNo " _
        & "WHERE [Membership status] = 'Active' AND [Door Access] = 'Disabled' and JSCACardNum <> 0)"
    Debug.Print strSQL
    DoCmd.RunSQL (strSQL)

    strSQL = "UPDATE tbl_Members " _
        & "SET [ProcessingStatus] = [ProcessingStatus] & 'Door Access not enabled.  Privileges NOT updated.  '" _
        & "WHERE [Membership status] = 'Active' AND [Door Access] = 'Disabled' and JSCACardNum <> 0"
    Debug.Print strSQL
    DoCmd.RunSQL (strSQL)

End Sub


Sub RemoveAllBUTJSCAandPending()

    Dim strSQL As String
    
    ' DoCmd.SetWarnings False
        
' Removes all privleges EXCEPT Pending Status (as we want to keep what they had before they became pending) and JSCA staff (GroupID 4)
    strSQL = "DELETE * FROM t_d_Privilege " _
        & "WHERE  t_d_Privilege.f_RecID NOT IN ( " _
        & "SELECT t_d_Privilege.f_RecID " _
        & "FROM tbl_Members RIGHT JOIN (t_b_Consumer INNER JOIN t_d_Privilege ON t_b_Consumer.f_ConsumerID = t_d_Privilege.f_ConsumerID) ON tbl_Members.[Member ID] = t_b_Consumer.f_WorkNo " _
        & "WHERE Left([Membership status],7) = 'Pending' OR f_GroupID = 4)"
    Debug.Print strSQL
    DoCmd.RunSQL (strSQL)

End Sub


Sub UpdatePendingRecordStatus()

    Dim strSQL As String
    
    ' DoCmd.SetWarnings False
        
' Update Processing Status for all 'Pending Records' as these are not processed
    strSQL = "UPDATE tbl_Members SET tbl_Members.ProcessingStatus = 'Pending Status - Did not process.  ' " _
        & "WHERE [Membership status] Like 'Pending*' "
    Debug.Print strSQL
    DoCmd.RunSQL (strSQL)

End Sub


Public Sub setPriv(MemberID As Long, doorName As String, ByVal priv As Boolean)

    Dim dbThis As Database
    Dim strSQL As String
    Dim rs As Recordset
    Dim ConsumerID As Long
    Dim DoorID As Long

    Set dbThis = CurrentDb

    ' First, get f_ConsumerID for this MemberID, so we install/delete the right person in the card system
      strSQL = " SELECT f_ConsumerID FROM t_b_Consumer WHERE f_WorkNo= '" & MemberID & "' ;"
    Set rs = dbThis.OpenRecordset(strSQL)
    If rs.EOF Then
'        MsgBox "ERROR: Can't find matching MemberID in cardDB! "
        Exit Sub
    Else
        ConsumerID = rs!f_ConsumerID
    End If
    
    ' Get the doorID for the requested doorName
    strSQL = " SELECT f_DoorID FROM t_b_Door WHERE f_DoorName='" & doorName & "';"
    Set rs = dbThis.OpenRecordset(strSQL)
    If rs.EOF Then
        MsgBox "ERROR: Can't find door name in cardDB! "
        Exit Sub
    Else
        DoorID = rs!f_DoorID
    End If
    
   ' See if the Consumer and doorID already have a clearance
   strSQL = " SELECT f_ConsumerID, f_DoorID FROM t_d_Privilege  " _
            & " WHERE ( (f_ConsumerID= " & ConsumerID & ") AND (f_DoorID= " & DoorID & ") );"
   Set rs = dbThis.OpenRecordset(strSQL)
   
   ' We want priv, and have none, so install
   If priv And rs.EOF Then
        With rs
            .AddNew
            !f_ConsumerID = ConsumerID
            !f_DoorID = DoorID
            .Update
            .Bookmark = rs.LastModified   ' Make current record the last modifed
            Debug.Print "ADD PRIV: f_ConsumerID = " & rs!f_ConsumerID & "  f_DoorID = "; rs!f_DoorID & " "
            ' log
        End With
    End If
   
   ' Want priv, and have, so do nothing
   
   ' Don't want priv, and have, so delete
    If Not priv And Not rs.EOF Then
        rs.Delete
        Debug.Print "DEL PRIV: ConsumerID = " & ConsumerID & "  DoorID = "; DoorID & " "
        ' log
    End If

   ' Don't want priv, and don't have, so do nothing
   
   rs.Close

End Sub


Function checkWindLevel3(MemberID As Long, rsThisMember As Recordset) As Boolean
    
    checkWindLevel3 = False
    
    If Not IsNull(rsThisMember![Windsurfing Level 3]) And rsThisMember![Windsurfing Level 3] <> "" Then
        checkWindLevel3 = True
    End If

End Function