Option Compare Database
Option Explicit
Private Sub cboFind_AfterUpdate()
On Error GoTo err_cboFind
    If Me![cboFind] <> "" Then
        Me.Filter = "[UnitNumber] = " & Me![cboFind]
        Me.FilterOn = True
    End If
Exit Sub
err_cboFind:
    MsgBox Err.Description
    Exit Sub
End Sub
Private Sub cmdAddNew_Click()
On Error GoTo err_cmdAddNew
    DoCmd.OpenForm "FRM_SkeletonDEscription", acNormal, , , acFormAdd
    Forms![FRM_SkeletonDEscription]![txtUnit] = Me![txtUnit]
    DoCmd.Close acForm, Me.Name
Exit Sub
err_cmdAddNew:
    MsgBox Err.Description
Exit Sub
End Sub
Private Sub cmdAddNewUnit_Click()
On Error GoTo err_cmdAddNewUnit
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    DoCmd.GoToControl "txtUnit"
Exit Sub
err_cmdAddNewUnit:
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
Private Sub CmdOpenAgeSexFrm_Click()
On Error GoTo Err_CmdOpenAgeSexFrm_Click
    Call DoRecordCheck("HR_ageing and sexing", Me![txtUnit], Me![txtIndivid], "Unit Number")
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
Private Sub Combo28_BeforeUpdate(Cancel As Integer)
End Sub
