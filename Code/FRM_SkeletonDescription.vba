Option Compare Database
Option Explicit
Private Sub cboFind_AfterUpdate()
On Error GoTo err_cboFind
    If Me![cboFind] <> "" Then
        Me.Filter = "[UnitNumber] = " & Me![cboFind] & " AND [Individual Number] = " & Me!cboFind.Column(1)
        Me.FilterOn = True
    End If
Exit Sub
err_cboFind:
    MsgBox Err.Description
    Exit Sub
End Sub
Private Sub cboFind_NotInList(NewData As String, response As Integer)
On Error GoTo err_cbofindNot
    MsgBox "This skeleton number does not exist in the database", vbInformation, "No Match"
    response = acDataErrContinue
    Me![cboFind].Undo
    DoCmd.GoToControl "CmdOpenUnitDescFrm"
Exit Sub
err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdAll_Click()
On Error GoTo err_all
    Me![cboFind].RowSource = "SELECT [HR_BasicSkeletonData].[UnitNumber], [HR_BasicSkeletonData].[Individual number] FROM HR_BasicSkeletonData ORDER BY [HR_BasicSkeletonData].[UnitNumber], [HR_BasicSkeletonData].[Individual number];"
Exit Sub
err_all:
    MsgBox Err.Description
    Exit Sub
End Sub
Private Sub cmdGuide_Click()
On Error GoTo err_cmdGuide
    DoCmd.OpenForm "frm_pop_skeletonguide", acNormal, , , acFormReadOnly
Exit Sub
err_cmdGuide:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdNewSkeleton_Click()
On Error GoTo err_cmdNew
    Dim thisUnit
    thisUnit = Me![txtUnit]
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    Me![txtUnit].Locked = False
    DoCmd.GoToControl "txtUnit"
    Me![txtUnit] = thisUnit
    Me![txtUnit].Locked = True
    DoCmd.GoToControl "txtIndivid"
Exit Sub
err_cmdNew:
    MsgBox Err.Description
    Exit Sub
End Sub
Private Sub CmdOpenAgeSexFrm_Click()
On Error GoTo Err_CmdOpenAgeSexFrm_Click
    Call DoRecordCheck("HR_ageing and sexing", Me![txtUnit], Me![txtIndivid], "Unit Number")
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "FRM_Ageing-sexing form"
    DoCmd.OpenForm stDocName, , , "[Unit Number] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    Forms![FRM_Ageing-sexing form]!cboFind.RowSource = "SELECT [HR_ageing and sexing].[unit number], [HR_ageing and sexing].[Individual number] FROM [HR_ageing and sexing] WHERE [HR_ageing and sexing].[Unit Number] = " & Me![UnitNumber] & " ORDER BY [HR_ageing and sexing].[Unit Number], [HR_ageing and sexing].[Individual number];"
    DoCmd.Close acForm, Me.Name
Exit_CmdOpenAgeSexFrm_Click:
    Exit Sub
Err_CmdOpenAgeSexFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenAgeSexFrm_Click
End Sub
Private Sub CmdOpenMainMenuFrm_Click()
On Error GoTo Err_CmdOpenMainMenuFrm_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "FRM_MAINMENU"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, Me.Name
Exit_CmdOpenMainMenuFrm_Click:
    Exit Sub
Err_CmdOpenMainMenuFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenMainMenuFrm_Click
End Sub
Private Sub CmdOpenUnitDescFrm_Click()
On Error GoTo Err_cmdUnitDesc_Click
If Me![txtUnit] <> "" Then
    Dim checknum, sql
    checknum = DLookup("[UnitNumber]", "[HR_UnitDescription]", "[UnitNumber] = " & Me![txtUnit])
    If IsNull(checknum) Then
        sql = "INSERT INTo [HR_UnitDescription] ([UnitNumber]) VALUES (" & Me![txtUnit] & ");"
        DoCmd.RunSQL sql
    End If
    DoCmd.OpenForm "Frm_UnitDescription", acNormal, , "[UnitNumber] = " & Me![txtUnit], acFormPropertySettings
    DoCmd.Close acForm, Me.Name
Else
    MsgBox "No Unit number is present, cannot open the Unit Description form", vbInformation, "No Unit Number"
End If
Exit Sub
Err_cmdUnitDesc_Click:
    MsgBox Err.Description
    Exit Sub
End Sub
Private Sub Form_Delete(Cancel As Integer)
On Error GoTo err_delete
Dim permiss
permiss = GetGeneralPermissions
If (permiss = "ADMIN") Then
    Dim response
    response = MsgBox("Deleting this skeleton will mean permanent deletion of all data associated with this particular skeleton in this database." & Chr(13) & Chr(13) & "Do you really want to delete " & Me![txtUnit] & ".B" & Me![txtIndivid] & "?", vbCritical + vbYesNo, "Critical Delete")
    If response = vbNo Then
        Cancel = True
    Else
        Cancel = False
    End If
Else
    MsgBox "You do not have permission to delete this record, please contact your team leader"
    Cancel = True
End If
Exit Sub
err_delete:
    Call General_Error_Trap
    Exit Sub
End Sub
