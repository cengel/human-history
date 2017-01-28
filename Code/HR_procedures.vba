Option Compare Database
Option Explicit
Sub ReturnToMenu(frm)
On Error GoTo err_menu
    DoCmd.OpenForm "FRM_MAINMENU"
    DoCmd.Close acForm, frm.Name
Exit Sub
err_menu:
    MsgBox Err.Description
    Exit Sub
End Sub
Sub DoRecordCheck(tblName, Unit, Individ, UnitfldName)
On Error GoTo err_check
    Dim mydb As DAO.Database, myrs As DAO.Recordset, sql As String, sql2 As String
    sql = "SELECT [" & UnitfldName & "], [Individual Number] FROM [" & tblName & "] WHERE [" & UnitfldName & "] = " & Unit & " AND [Individual Number] = " & Individ & ";"
    Set mydb = CurrentDb
    Set myrs = mydb.OpenRecordset(sql)
    If myrs.BOF And myrs.EOF Then
        sql2 = "INSERT INTO [" & tblName & "] ([" & UnitfldName & "], [Individual Number]) VALUES (" & Unit & "," & Individ & ");"
        DoCmd.RunSQL sql2
    End If
    myrs.Close
    Set myrs = Nothing
    mydb.Close
    Set mydb = Nothing
Exit Sub
err_check:
    MsgBox Err.Description
    Exit Sub
End Sub
Sub SortOutButtons(frm As Form)
On Error GoTo err_Sort
If frm![cboAgeCategory] <> "" Then
        If frm![cboAgeCategory] = 0 Then
            frm![CmdOpenNeonateFrm].Enabled = True
            frm![CmdOpenJuvenileFrm].Enabled = False
            frm![CmdOpenAdultFrm].Enabled = False
        ElseIf frm![cboAgeCategory] = 1 Or frm![cboAgeCategory] = 2 Or frm![cboAgeCategory] = 3 Then
            frm![CmdOpenNeonateFrm].Enabled = False
            frm![CmdOpenJuvenileFrm].Enabled = True
            frm![CmdOpenAdultFrm].Enabled = False
        ElseIf frm![cboAgeCategory] = 4 Or frm![cboAgeCategory] = 5 Or frm![cboAgeCategory] = 6 Or frm![cboAgeCategory] = 7 Then
            frm![CmdOpenNeonateFrm].Enabled = False
            frm![CmdOpenJuvenileFrm].Enabled = False
            frm![CmdOpenAdultFrm].Enabled = True
        Else
            frm![CmdOpenNeonateFrm].Enabled = True
            frm![CmdOpenJuvenileFrm].Enabled = True
            frm![CmdOpenAdultFrm].Enabled = True
        End If
   Else
        frm![CmdOpenNeonateFrm].Enabled = False
            frm![CmdOpenJuvenileFrm].Enabled = False
            frm![CmdOpenAdultFrm].Enabled = False
    End If
Exit Sub
err_Sort:
    MsgBox Err.Description
    Exit Sub
End Sub
