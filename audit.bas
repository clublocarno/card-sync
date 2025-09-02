Option Compare Database

Sub RunAudit()

'    LoadDatafromWA ("Full")
    DoCmd.OpenQuery "qry_Audit_MembersWITHOUTAccess"

End Sub