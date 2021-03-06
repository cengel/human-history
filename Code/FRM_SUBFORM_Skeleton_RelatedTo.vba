Option Compare Database
Option Explicit
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
Private Sub cmdMark_Click()
On Error GoTo err_cmdMark
    If (Me![Unit] = "" Or IsNull(Me![Unit])) Or (Me!IndividualNumber = "" Or IsNull(Me!IndividualNumber)) Then
        MsgBox "No relationships established yet so function not available", vbInformation, "No Relationships"
    Else
        DoCmd.OpenForm "FRM_SUBFORM_IncludeInAgeCategory", , , , , , "([HR_Skeleton_RelatedTo_Skeleton].Unit=" & Me!Unit & " AND [HR_Skeleton_RelatedTo_Skeleton].IndividualNumber=" & Me!IndividualNumber & ") OR ([HR_Skeleton_RelatedTo_Skeleton].RelatedToUnit=" & Me!Unit & " AND [HR_Skeleton_RelatedTo_Skeleton].RelatedToIndividualNumber=" & Me!IndividualNumber & ")"
    End If
Exit Sub
err_cmdMark:
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
