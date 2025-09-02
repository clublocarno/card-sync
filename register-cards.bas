Option Explicit
Option Compare Database


Sub GetJSCACardStatus()

    Dim strSQL As String

    ' DoCmd.SetWarnings False
    
' Fix issues where update can't happen from source table ... so create a copy of JSCA table
    
'    DoCmd.RunSQL "Delete From dbo_vw_get_LSC_new2015_copy"
'    DoCmd.OpenQuery "qry_CopyJSCATableLocally"

' Update JSCA Card Status based on following criteria:
'   Exact match on last 5 digits of card
'   Exact match on left (last name, 4)
'   Exact on first initial
    
    DoCmd.OpenQuery "qry_GetJSCAStatus"

'   OLD.  Discontinued March 2015
'   Exact match on last 5 digits
'   Exact match on last name
'   Exact on first initial

'    strSQL = "UPDATE tbl_Members INNER JOIN dbo_vw_get_LSC_Affiliated ON " _
'        & "(tbl_Members.JSCACardNum = dbo_vw_get_LSC_Affiliated.CARDNUM) AND " _
'        & "(tbl_Members.[Last name] = dbo_vw_get_LSC_Affiliated.LAST_NAME) " _
'        & "SET tbl_Members.JSCACardStatus = [STATUS], tbl_Members.FAMNUM = dbo_vw_get_LSC_Affiliated.FAMNUM " _
'        & "WHERE Left([First name], 1) = Left([FIRST_NAME], 1)"
'    DoCmd.RunSQL (strSQL)
    

End Sub


Sub EnsureJSCACardIsRegisteredinCardSystem()

On Error GoTo Err_Process
    
    Dim dbThis As Database
    Dim rsMembers As Recordset
    Dim strSQL As String
    Dim vMemberID As Long
    Dim vFullName As String
    Dim vCardNumFromMember As String
    Dim vCardNumInDB As String
    Dim ErrorDescr As String

    Set dbThis = CurrentDb
    
    AddLeadingZeros  ' to tbl_Members.JSCACardNum_Text as it is used in the following expression
    
' In order to avoid enabling/diabling fleets before Apr 1 (too early),
' form Jan 1 - Mar 32 each year, DO NOT use the Card Software's UPDATE function,
' (which delivers the settings to the door hardware)
' However, it's to ?? sync during Jan1-Mar31, (as you get members joining) to get their cards registered.

       
    
    ' DoCmd.SetWarnings False
    
' Update tbl_Members with flag indicating cards that are already registered
' Card Status Desc = 0 indicates active card
    strSQL = "UPDATE tbl_Members INNER JOIN (t_b_Consumer INNER JOIN t_b_IDCard ON t_b_Consumer.f_ConsumerID = t_b_IDCard.f_ConsumerID) ON tbl_Members.[Member ID] = t_b_Consumer.f_WorkNo " _
        & "SET tbl_Members.CardRegistered = True " _
        & "WHERE t_b_IDCard.f_CardStatusDesc = '0' " _
        & "AND tbl_Members.JSCACardNum_Text = Right(f_CardNO,5)"
'    Debug.Print strsql
    DoCmd.RunSQL (strSQL)
  

' Select ONLY Cards that need to be registered - Active Locarno and JSCA Members
    strSQL = "SELECT * FROM tbl_Members WHERE " _
        & "[Membership status] = 'Active' " _
        & "AND [JSCACardStatus] = 'Valid' " _
        & "AND CardRegistered = False " _
        & "AND JSCACardNum <> 0 " _
        & "AND [Door Access] = 'Enabled'"
    Set rsMembers = dbThis.OpenRecordset(strSQL)
    
    If rsMembers.RecordCount = 0 Then
        Exit Sub
    End If
        
' Walk through current members
    Do While Not rsMembers.EOF       ' avoids MoveFirst, as will choke for no records
    
        vMemberID = rsMembers![Member ID]
        vFullName = rsMembers![First name] & " " & rsMembers![Last name]

' Update or Add Member is in the Card System
        UpdateAddMemberInCardSystem vMemberID, vFullName


' Grad FAMNUM from JSCA table and Concate with number provided by Member
        vCardNumFromMember = rsMembers![FAMNUM] & rsMembers![JSCACardNum_Text]
        
        vCardNumInDB = getCard(vMemberID)
        
        If vCardNumInDB = "" Then
            RegisterCard vMemberID, vCardNumFromMember
            EmailMember vMemberID, "Card has been registered"
            With rsMembers
                .Edit
                !CardRegistationStatus = "Registered New Card"
                !CardRegistered = True
                !EmailSent = "Email Sent"
                .Update
            End With
        Else
' Returns a number, which is not the same number, based on the original query to start to sub
            InactiveCard vMemberID, vCardNumInDB
            RegisterCard vMemberID, vCardNumFromMember
            EmailMember vMemberID, "New Card has been registered"
            With rsMembers
                .Edit
                !CardRegistationStatus = "Registered New Card and inactived Old Card"
                !CardRegistered = True
                !EmailSent = "Email Sent"
                .Update
            End With
        End If
        
        rsMembers.MoveNext
    
    Loop

Exit_Process:
    Exit Sub

Err_Process:
    Debug.Print Err.Description
    ErrorDescr = Replace(Err.Description, "'", " ")
    strSQL = "INSERT INTO tbl_ErrorLog (ProcessStep, Description) VALUES ('Sub EnsureJSCACardIsRegistered', '" & ErrorDescr & "')"
    DoCmd.RunSQL (strSQL)
        Resume Exit_Process
        
End Sub


Sub UpdateAddMemberInCardSystem(vMemberID As Long, vFullName As String)

    Dim strSQLMemberID As String
    Dim ConsumerID As Long
    Dim ConsumerNO As Long
    Dim dbThis As Database
    Dim rs As Recordset
    Dim rsAdd As Recordset
    Dim ThisYear As Integer
    Dim DatePaidCurYear As Date
    
    Set dbThis = CurrentDb

' Check/set MemberID in the cDB
    strSQLMemberID = "SELECT * FROM t_b_Consumer WHERE t_b_Consumer.f_WorkNo = '" & vMemberID & "';"
    Set rs = dbThis.OpenRecordset(strSQLMemberID)
    
    If rs.EOF Then      ' no matching record(s), so install name and MembershipID
        ' First, determine next available ConsumerNO
        Set rs = dbThis.OpenRecordset("SELECT Max(f_ConsumerNO) as maxConsumerNO FROM t_b_Consumer ;")
        If IsNull(rs!maxConsumerNO) Then        ' cater for a complete rebuild from scratch
            ConsumerNO = 1
        Else
            ConsumerNO = rs!maxConsumerNO + 1
        End If
        
       ' Append new personel record
        Set rsAdd = dbThis.OpenRecordset("t_b_Consumer")       ' open table recordset
        With rsAdd
            .AddNew
            !f_ConsumerNO = ConsumerNO
            !f_ConsumerName = vFullName
            !f_WorkNo = vMemberID
            .Update                       ' Note: After update rs pts to prior rs spot, use Bookmark as in next line
            .Bookmark = rs.LastModified   ' Make current record the last modifed
'            ConsumerID = !f_ConsumerID     ' get a copy of this for possible card registration below
' WRONG            Debug.Print "ADDED: f_ConsumerID (auto) = " & !f_ConsumerID & "  f_ConsumerNO = "; !f_ConsumerNO & " f_ConsumerName = " & !f_ConsumerName & " f_WorkNo = " & !f_WorkNo & ""
        End With
    
    Else ' Member in database, thus update.
        
' Check/set name (syncs name changes from mDB )
        If rs!f_ConsumerName <> vFullName Then    ' need to update name in cDB
            With rs
                .Edit
                !f_ConsumerName = vFullName
                .Update
            End With
            ' log
        End If
            
    End If
    
End Sub

Function getCard(MemberID As Long) As String

' attempts to get  the card number for a memberID
Dim dbThis As Database
Dim strSQL As String
Dim rs As Recordset

   Set dbThis = CurrentDb
   ' get the cardnumber for this memberID, ensure not flagged as Lost
    strSQL = "SELECT f_CardNO FROM  t_b_IDCard INNER JOIN t_b_Consumer " _
       & " ON t_b_IDCard.f_ConsumerID = t_b_Consumer.f_ConsumerID " _
       & " WHERE ( (t_b_Consumer.f_WorkNo= '" & MemberID & "' ) " _
       & " AND (t_b_IDCard.f_CardStatusDesc='0') );"
            
    Set rs = dbThis.OpenRecordset(strSQL)
    If rs.EOF Then  ' no card
        getCard = ""
    Else
        getCard = rs!f_CardNO
    End If

    rs.Close

End Function


Sub RegisterCard(vMemberID As Long, CardNO As String)

Dim ConsumerID As Long
Dim dbThis As Database
Dim rs As Recordset
    
Set dbThis = CurrentDb
    

    ConsumerID = DLookup("f_ConsumerID", "t_b_Consumer", "f_WorkNo='" & vMemberID & "'")
    
    Set rs = dbThis.OpenRecordset("t_b_IDCard")       ' open table recordset
    With rs
        .AddNew
        !f_CardNO = CardNO
        !f_ConsumerID = ConsumerID
        .Update                       ' Note: After update rs pts to prior rs spot, use Bookmark as in next line
        .Bookmark = rs.LastModified   ' Make current record the last modifed
        Debug.Print "ADDED: f_CardID (auto) = " & !f_CardID & "  f_CardNO = "; !f_CardNO & " f_ConsumerID = " & !f_ConsumerID & ""
    End With

End Sub


Sub InactiveCard(vMemberID As Long, CardNO As String)

Dim ConsumerID As Long
Dim dbThis As Database
Dim rs As Recordset
Dim strSQL As String
    
Set dbThis = CurrentDb
    

    ConsumerID = DLookup("f_ConsumerID", "t_b_Consumer", "f_WorkNo='" & vMemberID & "'")
    
    strSQL = "SELECT * FROM t_b_IDCard WHERE t_b_IDCard.f_CardNO = '" & CardNO & "'"
    Set rs = dbThis.OpenRecordset(strSQL)
    With rs
        .Edit
        !f_CardStatusDesc = "Lost"
        .Update
        Debug.Print "Inactivated: f_CardID = " & !f_CardID & "  f_CardNO = "; !f_CardNO & " f_ConsumerID = " & !f_ConsumerID & ""
    End With

End Sub


Sub AddLeadingZeros()

    Dim strSQL As String
    
    ' DoCmd.SetWarnings False
    
    strSQL = "UPDATE tbl_Members " _
        & "SET tbl_Members.JSCACardNum_Text  = '0' & tbl_Members.JSCACardNum_Text " _
        & "WHERE len(tbl_Members.JSCACardNum_Text) = 4 "
    DoCmd.RunSQL (strSQL)

    strSQL = "UPDATE tbl_Members " _
        & "SET tbl_Members.JSCACardNum_Text  = '00' & tbl_Members.JSCACardNum_Text " _
        & "WHERE len(tbl_Members.JSCACardNum_Text) = 3 "
    DoCmd.RunSQL (strSQL)

    strSQL = "UPDATE tbl_Members " _
        & "SET tbl_Members.JSCACardNum_Text  = '000' & tbl_Members.JSCACardNum_Text " _
        & "WHERE len(tbl_Members.JSCACardNum_Text) = 2 "
    DoCmd.RunSQL (strSQL)


End Sub