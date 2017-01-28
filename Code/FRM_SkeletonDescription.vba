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
Private Sub cboFind_NotInList(NewData As String, Response As Integer)
On Error GoTo err_cbofindNot
    MsgBox "This skeleton number does not exist in the database", vbInformation, "No Match"
    Response = acDataErrContinue
    Me![cboFind].Undo
    DoCmd.GoToControl "CmdOpenUnitDescFrm"
Exit Sub
err_cbofindNot:
    Call General_Error_Trap
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
