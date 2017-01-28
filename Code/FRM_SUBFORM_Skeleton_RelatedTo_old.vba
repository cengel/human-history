Option Compare Database
Private Sub cmdAddRelation_Click()
On Error GoTo err_skel
    Dim strArgs
    strArgs = Forms![FRM_SkeletonDescription]![txtUnit] & "." & Forms![FRM_SkeletonDescription]![txtIndivid]
    DoCmd.OpenForm "FRM_pop_Add_Skel_Relation", acNormal, , , acFormPropertySettings, acDialog, strArgs
    Me.Requery
Exit Sub
err_skel:
    MsgBox Err.Description
    Exit Sub
End Sub
