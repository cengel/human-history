Option Compare Database
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
Private Sub cmdAll_Click()
On Error GoTo err_all
    Me.FilterOn = False
    Me.Filter = ""
Exit Sub
err_all:
    MsgBox Err.Description
    Exit Sub
End Sub
Private Sub cmdGuide_Click()
On Error GoTo err_cmdGuide
    DoCmd.OpenForm "frm_pop_tooth_guide", acNormal, , , acFormReadOnly
Exit Sub
err_cmdGuide:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub CmdOpenDecidTeethFrm_Click()
On Error GoTo Err_CmdOpenDecidTeethFrm_Click
        Dim stDocName As String
        Dim stLinkCriteria As String
        stDocName = "FRM_PERMANENT_TEETH"
        DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
        DoCmd.Close acForm, Me.Name
Exit_CmdOpenDecidTeethFrm_Click:
    Exit Sub
Err_CmdOpenDecidTeethFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenDecidTeethFrm_Click
End Sub
Private Sub CmdOpenJuvenileFrm_Click()
On Error GoTo Err_CmdOpenJuvFrm_Click
    Call DoRecordCheck("HR_Juvenile_Cranial_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Juvenile_shoulder_hip", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Juvenile_axial", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Juvenile_Arm_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Juvenile_Leg_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "FRM_Juvenile"
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name
Exit_CmdOpenJuvFrm_Click:
    Exit Sub
Err_CmdOpenJuvFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenJuvFrm_Click
End Sub
Private Sub CmdOpenAdultFrm_Click()
On Error GoTo Err_CmdOpenAdultFrm_Click
    Call DoRecordCheck("HR_Adult_Cranial_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Adult_shoulder_hip", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Adult_Axial_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Adult_Arm_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Call DoRecordCheck("HR_Adult_Leg_Data", Me![txtUnit], Me![txtIndivid], "UnitNumber")
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "FRM_Adult"
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name
Exit_CmdOpenAdultFrm_Click:
    Exit Sub
Err_CmdOpenAdultFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenAdultFrm_Click
End Sub
Private Sub CmdOpenAgeSexFrm_Click()
On Error GoTo Err_CmdOpenAgeSexFrm_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "FRM_Ageing-sexing form"
    DoCmd.OpenForm stDocName, , , "[Unit Number] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name
Exit_CmdOpenAgeSexFrm_Click:
    Exit Sub
Err_CmdOpenAgeSexFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenAgeSexFrm_Click
End Sub
Private Sub CmdOpenMainMenuFrm_Click()
Call ReturnToMenu(Me)
End Sub
Private Sub Command462_Click()
On Error GoTo Err_CmdOpenUnitDescFrm_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "FRM_SkeletonDescription"
    DoCmd.OpenForm stDocName, , , "[UnitNumber] = " & Me![txtUnit] & " AND [Individual Number] = " & Me![txtIndivid]
    DoCmd.Close acForm, Me.Name
Exit_CmdOpenUnitDescFrm_Click:
    Exit Sub
Err_CmdOpenUnitDescFrm_Click:
    MsgBox Err.Description
    Resume Exit_CmdOpenUnitDescFrm_Click
End Sub
Private Sub Form_Current()
On Error GoTo err_current
Exit Sub
err_current:
    General_Error_Trap
    Exit Sub
End Sub
