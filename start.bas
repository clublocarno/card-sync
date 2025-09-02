Option Compare Database

Public Sub UpdateStatusBar(vMEssage As String)

    Forms![fmHomePage]![txtStatus].Value = Time() & ": " & vMEssage & vbNewLine & Forms![fmHomePage]![txtStatus].Value

End Sub


Public Sub ExecuteProcess(vScope As String)

     DoCmd.SetWarnings False
        
' My card number: 47091

' Run Wild Apricot API to load data into tbl_Members - Module: 1 WildApricot API
    Debug.Print "Starting Load Process"
    UpdateStatusBar ("Commencing API request")
    
    GetDatafromWP (vScope)

    If vScope = "FullLoad" Then
' Remove all Privleges except Pending status and JSCA staff (latter which are managed through the Card system).
        Debug.Print "Removing all Privileges Except ..."
        UpdateStatusBar ("Removing all Privileges Except ...")

        RemoveAllBUTJSCAandPending
' Remove all Members to Reprocess
        DoCmd.RunSQL ("DELETE tbl_Members_ToReprocess.* FROM tbl_Members_ToReprocess ")

    End If
        
' If members have problems with their JSCA information, an email is sent to them asking them to
' contact the JSCA admin to update their information.  When they do so, this will either result in
' updated information on Wild Apricot or "dbo_vw_get_LSC_Affiliated".
' The LoadDatafromWA procedure will get changes in Wild Apricot
' This procedure will add these members back into the process, pulling from a
' temporary table that is populated at the end of the entire process.
' Add members from "tbl_Members_ToReprocess" to "tbl_Members"
    AddMembersToReprocess

' Wild Apricot does not store historical information once a membership is changed to Pending-Renewal or Pending-Level Change
' Thus in order to avoid changing a member's priviledges until a transition is complete, there record is no processed,
' thus leaving their information intact.
    UpdatePendingRecordStatus



' *********

' If have connectivity to the JSCA Server, check records.  Else, skip.
'If 1 = 0 Then  ' Bypass Validation
If 1 = 1 Then  ' Validate Cards

'    DoCmd.RunSQL ("Delete * from dbo_vw_get_LSC_new2015_local")
'    DoCmd.OpenQuery "qry_Copy_JSCA_Table_local"
'    DoCmd.OpenQuery "qry_GetJSCAStatus_from_Local"

    ' Run process to match members with JSCA view from their membership database
        UpdateStatusBar ("Identifing Valid JSCA Cards")
        GetJSCACardStatus
    
    ' Run Process to Ensure Card is Registered. Module 2 Register Cards
        Debug.Print "Starting Ensuring JSCA Card is Registered"
        UpdateStatusBar ("Starting Ensuring JSCA Card is Registered")
    
Else

    DoCmd.OpenQuery "qry: Bypass_JSCA_Server"

End If

    EnsureJSCACardIsRegisteredinCardSystem

' Update Card System - Alan's original code to update privileges.  Module: 3 Updates to Card System
    Debug.Print "Starting Updating Privleges"
    UpdateStatusBar ("Starting Updating Privleges")
    UpdateCardSystem
    
' Module: 4 Emails
'If 1 = 0 Then  ' Bypass Emails
If 1 = 1 Then  ' Send Emails
    Debug.Print "Starting Procesing Based on JSCA Card"
    UpdateStatusBar ("Starting Email Members With Problems")
    EmailMembersWithProblems

' If members have problems with their JSCA information, an email is sent to them asking them to
' contact the JSCA admin to update their information.  When they do so, this information will
' appear in "dbo_vw_get_LSC_Affiliated".
' Since the trigger to process individuals is a change in Wild Apricot, this process saves
' these members for reprocessing later.
    SaveMembersToReprocess
End If

' Email status
    Debug.Print "Email Status"
    UpdateStatusBar ("Email Status")
    EmailStatus
    
' Keep Audit Log
    ' DoCmd.SetWarnings False
    DoCmd.OpenQuery ("qry_AddToAudit")


    DoCmd.SetWarnings True

' TO DOS


' Original Error Code
' Err_btn_sync_Click
' Exit_btn_sync_Click

End Sub


Sub SaveMembersToReprocess()

    ' DoCmd.SetWarnings False
    
    DoCmd.RunSQL ("DELETE tbl_Members_ToReprocess.* FROM tbl_Members_ToReprocess ")
    DoCmd.OpenQuery ("qry_Members_ToReprocess_Add")

End Sub

Sub AddMembersToReprocess()

    Dim strSQL As String
    
    ' DoCmd.SetWarnings False
    
' Remove records that are already in tbl_Members
    strSQL = "DELETE tbl_Members_ToReprocess.* FROM tbl_Members_ToReprocess " _
        & "WHERE tbl_Members_ToReprocess.[Member ID] IN " _
        & "(SELECT tbl_Members.[Member ID] FROM tbl_Members)"
    DoCmd.RunSQL strSQL

' Add records to tbl_Members
    DoCmd.OpenQuery ("qry_Members_ToReprocess_AddBack")
    

End Sub