Option Compare Database
Option Explicit
Private Sub cmdAddRelation_Click()
On Error GoTo err_skel
    Dim getXFind, getNotes, sql
    getXFind = InputBox("Enter the X find number that this individual number relates to:", "X Find Number Required")
    If getXFind <> "" Then
        getNotes = InputBox("Enter any notes or comments about this relationship:", "Notes or Comments")
        If getNotes <> "" Then
            sql = "INSERT INTO [HR_Skeleton_RelatedTo_XFind] ([Unit], [IndividualNumber], [XfindNumber], [Notes]) VALUES (" & Forms![FRM_SkeletonDescription]![txtUnit] & ", " & Forms![FRM_SkeletonDescription]![txtIndivid] & ", " & getXFind & ", '" & getNotes & "');"
        Else
            sql = "INSERT INTO [HR_Skeleton_RelatedTo_XFind] ([Unit], [IndividualNumber], [XfindNumber]) VALUES (" & Forms![FRM_SkeletonDescription]![txtUnit] & ", " & Forms![FRM_SkeletonDescription]![txtIndivid] & ", " & getXFind & ");"
        End If
        DoCmd.RunSQL sql
        Me.Requery
    End If
    DoCmd.GoToControl "cmdAddRelation"
Exit Sub
err_skel:
    MsgBox Err.Description
    Exit Sub
End Sub
Private Sub cmdDelete_Click()
On Error GoTo err_cmdDelete
    Dim resp
    resp = MsgBox("Do you really want to delete the relationship between skeleton " & Me![Unit] & ".B" & Me![IndividualNumber] & " and X find number" & Me![Unit] & ".X" & Me![XFindNumber] & "?", vbCritical + vbYesNo, "Confirm Deletion")
    If resp = vbYes Then
        Dim sql
        sql = "Delete FROM [HR_Skeleton_RelatedTo_XFind] WHERE [Unit] = " & Me![Unit] & " AND [IndividualNumber] = " & Me![IndividualNumber] & " AND [XFindNumber] = " & Me![XFindNumber] & ";"
        DoCmd.RunSQL sql
        Me.Requery
        DoCmd.GoToControl "cmdAddRelation"
    End If
Exit Sub
err_cmdDelete:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_open
Dim permiss
    permiss = GetGeneralPermissions
    If (permiss = "ADMIN") Then
        Me![cmdDelete].Enabled = True
    Else
        Me![cmdDelete].Enabled = False
    End If
Exit Sub
err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
