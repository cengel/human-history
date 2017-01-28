Option Compare Database
Private Sub cmdAddRelation_Click()
On Error GoTo err_skel
    Dim strArgs
    strArgs = Forms![FRM_SkeletonDescription]![txtUnit] & "." & Forms![FRM_SkeletonDescription]![txtIndivid]
    DoCmd.OpenForm "FRM_pop_Add_Skel_Relation", acNormal, , , acFormPropertySettings, acDialog, strArgs
    Me.Requery
    DoCmd.GoToControl "cmdAddRelation"
Exit Sub
err_skel:
    MsgBox Err.Description
    Exit Sub
End Sub
Private Sub cmdDelete_Click()
On Error GoTo err_cmdDelete
    Dim resp
    resp = MsgBox("Do you really want to delete the relationship between skeleton " & Me![Unit] & ".B" & Me![IndividualNumber] & " and " & Me![RelatedToUnit] & ".B" & Me![RelatedToIndividualNumber] & "?", vbCritical + vbYesNo, "Confirm Deletion")
    If resp = vbYes Then
        Dim sql
        sql = "Delete FROM [HR_Skeleton_RelatedTo_Skeleton] WHERE [Unit] = " & Me![Unit] & " AND [IndividualNumber] = " & Me![IndividualNumber] & " AND [RelatedToUnit] = " & Me![RelatedToUnit] & " AND [RelatedToIndividualNumber] = " & Me![RelatedToIndividualNumber] & ";"
        DoCmd.RunSQL sql
        sql = "Delete FROM [HR_Skeleton_RelatedTo_Skeleton] WHERE [Unit] = " & Me![RelatedToUnit] & " AND [IndividualNumber] = " & Me![RelatedToIndividualNumber] & " AND [RelatedToUnit] = " & Me![Unit] & " AND [RelatedToIndividualNumber] = " & Me![IndividualNumber] & ";"
        DoCmd.RunSQL sql
        Me.Requery
        DoCmd.GoToControl "cmdAddRelation"
    End If
Exit Sub
err_cmdDelete:
    Call General_Error_Trap
    Exit Sub
End Sub
