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
Private Sub cboFind_NotInList(NewData As String, Response As Integer)
On Error GoTo err_cbofindNot
    MsgBox "This unit number has not been entered yet", vbInformation, "No Match"
    Response = acDataErrContinue
    Me![cboFind].Undo
    DoCmd.GoToControl "cmdAddNewUnit"
Exit Sub
err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdAddNew_Click()
On Error GoTo err_cmdAddNew
Dim sql
    If Me![txtUnit] <> "" Then
        Dim checknum
        checknum = DLookup("[UnitNumber]", "[HR_BasicSkeletonData]", "[UnitNumber] = " & Me![txtUnit])
        If IsNull(checknum) Then
            sql = "INSERT INTO [HR_BasicSkeletonData] ([UnitNumber], [Individual Number]) VALUES (" & Me![txtUnit] & ", 1);"
            DoCmd.RunSQL sql
            Me.Refresh
            MsgBox "Individual number 1 added for Unit " & Me![txtUnit], vbInformation, "Record added"
        Else
            Dim mydb As Database, myrs As DAO.Recordset, lastnumber, nextnumber
            Set mydb = CurrentDb()
            sql = "SELECT HR_BasicSkeletonData.UnitNumber, HR_BasicSkeletonData.[Individual number] FROM HR_BasicSkeletonData WHERE HR_BasicSkeletonData.UnitNumber = " & Me![txtUnit] & " ORDER BY HR_BasicSkeletonData.UnitNumber, HR_BasicSkeletonData.[Individual number];"
            Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)
            If Not myrs.BOF And Not myrs.EOF Then
                myrs.MoveLast
                lastnumber = myrs![Individual number]
                nextnumber = lastnumber + 1
                sql = "INSERT INTO [HR_BasicSkeletonData] ([UnitNumber], [Individual Number]) VALUES (" & Me![txtUnit] & ", " & nextnumber & ");"
                DoCmd.RunSQL sql
                Me.Refresh
                MsgBox "Individual number " & nextnumber & " added for Unit " & Me![txtUnit], vbInformation, "Record added"
            Else
                 sql = "INSERT INTO [HR_BasicSkeletonData] ([UnitNumber], [Individual Number]) VALUES (" & Me![txtUnit] & ", 1);"
                DoCmd.RunSQL sql
                Me.Refresh
            End If
        End If
    Else
        MsgBox "You must enter a unit number first", vbInformation, "Unit Number Missing"
    End If
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
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_open
Exit Sub
err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub txtUnit_AfterUpdate()
On Error GoTo err_txtUnit
Exit Sub
err_txtUnit:
    Call General_Error_Trap
    Exit Sub
End Sub
