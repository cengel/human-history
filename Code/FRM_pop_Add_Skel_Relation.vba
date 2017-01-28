Option Compare Database
Option Explicit
Private Sub cboSelect_AfterUpdate()
End Sub
Private Sub cmdCancel_Click()
On Error GoTo err_cmdCancel
    DoCmd.Close acForm, Me.Name
Exit Sub
err_cmdCancel:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdOK_Click()
On Error GoTo err_cmdOK
Dim sql, sql1
Dim mydb As DAO.Database, myrs As DAO.Recordset
    If Me!cboSelect <> "" Then
        sql = "SELECT HR_Skeleton_RelatedTo_Skeleton.Unit, HR_Skeleton_RelatedTo_Skeleton.IndividualNumber, "
        sql = sql & "HR_Skeleton_RelatedTo_Skeleton.RelatedToUnit, HR_Skeleton_RelatedTo_Skeleton.RelatedToIndividualNumber "
        sql = sql & "FROM HR_Skeleton_RelatedTo_Skeleton "
        sql = sql & "WHERE ((HR_Skeleton_RelatedTo_Skeleton.Unit=" & Me!txtUnit & " AND HR_Skeleton_RelatedTo_Skeleton.IndividualNumber=" & Me!txtIndivid & ")"
        sql = sql & " AND "
        sql = sql & "(HR_Skeleton_RelatedTo_Skeleton.RelatedToUnit=" & Me!cboSelect.Column(1) & " AND HR_Skeleton_RelatedTo_Skeleton.RelatedToIndividualNumber=" & Me!cboSelect.Column(2) & "))"
        sql = sql & " OR"
        sql = sql & "((HR_Skeleton_RelatedTo_Skeleton.Unit=" & Me!cboSelect.Column(1) & " AND HR_Skeleton_RelatedTo_Skeleton.IndividualNumber=" & Me!cboSelect.Column(2) & ")"
        sql = sql & " AND "
        sql = sql & "(HR_Skeleton_RelatedTo_Skeleton.RelatedToUnit=" & Me!txtUnit & " AND HR_Skeleton_RelatedTo_Skeleton.RelatedToIndividualNumber=" & Me!txtIndivid & "));"
        Set mydb = CurrentDb
        Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)
        If Not (myrs.BOF And myrs.EOF) Then
            myrs.MoveLast
            If myrs.RecordCount = 1 Then
               myrs.MoveFirst
               If myrs![Unit] = CInt(Me!txtUnit) And myrs![IndividualNumber] = CInt(Me!txtIndivid) Then
                    sql = "INSERT INTO HR_Skeleton_RelatedTo_Skeleton (Unit, IndividualNumber, RelatedToUnit, RelatedToIndividualNumber) VALUES (" & Me![cboSelect].Column(1) & ", " & Me![cboSelect].Column(2) & ", " & Me![txtUnit] & ", " & Me![txtIndivid] & ");"
                    DoCmd.RunSQL sql
                ElseIf myrs![RelatedToUnit] = CInt(Me![txtUnit]) And myrs![RelatedToIndividualNumber] = CInt(Me!txtIndivid) Then
                    sql = "INSERT INTO HR_Skeleton_RelatedTo_Skeleton (Unit, IndividualNumber, RelatedToUnit, RelatedToIndividualNumber, Notes) VALUES (" & Me!txtUnit & ", " & Me![txtIndivid] & ", " & Me![cboSelect].Column(1) & ", " & Me![cboSelect].Column(2) & ", '" & Me!txtNotes & "');"
                    DoCmd.RunSQL sql
                End If
                MsgBox "Skeleton " & Me!txtUnit & ".B" & Me!txtIndivid & " was already related to Skeleton " & Me![cboSelect].Column(1) & ".B" & Me![cboSelect].Column(2) & " but this was not shown on screen, this problem has been recitfied. Press Cancel to exit this screen.", vbExclamation, "Relationship already exists"
            Else
                MsgBox "Skeleton " & Me!txtUnit & ".B" & Me!txtIndivid & " is already related to Skeleton " & Me![cboSelect].Column(1) & ".B" & Me![cboSelect].Column(2) & Chr(13) & Chr(13) & "Please choose another skeleton or press Cancel to exit this screen.", vbExclamation, "Relationship already exists"
            End If
        Else
            Dim Notes
            If Not IsNull(Me!txtNotes) Then
                Notes = Replace(Me!txtNotes, "'", "''")
            Else
                Notes = ""
            End If
            Dim OtherRelatedToUnit, OtherRelatedToIndivid, present
            sql1 = "SELECT * FROM HR_Skeleton_RelatedTo_Skeleton " & _
                    "WHERE HR_Skeleton_RelatedTo_Skeleton.Unit= " & Me![cboSelect].Column(1) & " AND HR_Skeleton_RelatedTo_Skeleton.IndividualNumber=" & Me![cboSelect].Column(2) & ";"
            Set mydb = CurrentDb
            Set myrs = mydb.OpenRecordset(sql1, dbOpenSnapshot)
            If Not (myrs.BOF And myrs.EOF) Then
                myrs.MoveFirst
                Do Until myrs.EOF
                    OtherRelatedToUnit = myrs![RelatedToUnit]
                    OtherRelatedToIndivid = myrs![RelatedToIndividualNumber]
                    present = DCount("[Unit]", "[HR_Skeleton_RelatedTo_Skeleton]", "[Unit] = " & Me![txtUnit] & " AND [IndividualNumber] = " & Me![txtIndivid] & " AND [RelatedToUnit] = " & OtherRelatedToUnit & " AND [RelatedToIndividualNumber] = " & OtherRelatedToIndivid)
                    If present = 0 Or IsNull(present) Then
                        MsgBox Me![cboSelect].Column(1) & ".B" & Me![cboSelect].Column(2) & " is in turn related to " & OtherRelatedToUnit & ".B" & OtherRelatedToIndivid & " and so this relationship will also exist here", vbInformation, "Relationship cascade"
                        sql = "INSERT INTO HR_Skeleton_RelatedTo_Skeleton (Unit, IndividualNumber, RelatedToUnit, RelatedToIndividualNumber, Notes) VALUES (" & Me!txtUnit & ", " & Me![txtIndivid] & ", " & OtherRelatedToUnit & ", " & OtherRelatedToIndivid & ", '" & myrs!Notes & "');"
                        DoCmd.RunSQL sql
                        sql = "INSERT INTO HR_Skeleton_RelatedTo_Skeleton (Unit, IndividualNumber, RelatedToUnit, RelatedToIndividualNumber, Notes) VALUES (" & OtherRelatedToUnit & ", " & OtherRelatedToIndivid & ", " & Me![txtUnit] & ", " & Me![txtIndivid] & ", '" & Notes & "');"
                        DoCmd.RunSQL sql
                    End If
                myrs.MoveNext
                Loop
            End If
            sql = "INSERT INTO HR_Skeleton_RelatedTo_Skeleton (Unit, IndividualNumber, RelatedToUnit, RelatedToIndividualNumber, Notes) VALUES (" & Me!txtUnit & ", " & Me![txtIndivid] & ", " & Me![cboSelect].Column(1) & ", " & Me![cboSelect].Column(2) & ", '" & Notes & "');"
            DoCmd.RunSQL sql
            sql = "INSERT INTO HR_Skeleton_RelatedTo_Skeleton (Unit, IndividualNumber, RelatedToUnit, RelatedToIndividualNumber, Notes) VALUES (" & Me![cboSelect].Column(1) & ", " & Me![cboSelect].Column(2) & ", " & Me![txtUnit] & ", " & Me![txtIndivid] & ", '" & "Relationship made from " & Me!txtUnit & "." & Me![txtIndivid] & ". " & Notes & "');"
            DoCmd.RunSQL sql
            DoCmd.Close acForm, Me.Name
        End If
        myrs.Close
        Set myrs = Nothing
        mydb.Close
        Set mydb = Nothing
    Else
        MsgBox "You must select an individual to relate to", vbInformation, "Action Cancelled"
    End If
Exit Sub
err_cmdOK:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_open
    If Not IsNull(Me.OpenArgs) Then
        Dim strArgs, unitnum, skelnum, dot
        strArgs = Me.OpenArgs
        dot = InStr(strArgs, ".")
        If dot > 0 Then
            unitnum = Left(strArgs, dot - 1)
            Me!txtUnit = unitnum
            skelnum = right(strArgs, Len(strArgs) - dot)
            Me!txtIndivid = skelnum
            Me![txtTitle] = "Relate Skeleton " & Me!txtUnit & ".B" & Me!txtIndivid & " to another Skeleton"
        Else
            MsgBox "Invalid identifier passed into the form, it must be the unit number and individual number"
            DoCmd.Close acForm, Me.Name
        End If
    End If
Exit Sub
err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
