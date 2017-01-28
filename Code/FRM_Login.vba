Option Compare Database
Private Sub cmdCancel_Click()
On Error GoTo cmdCancel_Click
Dim retval
retval = MsgBox("The system cannot continue without a login name and password." & Chr(13) & Chr(13) & "Are you sure you want to quit the system?", vbCritical + vbYesNo, "Confirm System Closure")
    If retval = vbYes Then
        MsgBox "The system will now quit", vbCritical + vbOKOnly, "Invalid Login"
        DoCmd.Quit acQuitSaveAll
    End If
    DoCmd.GoToControl "txtLogin"
Exit Sub
cmdCancel_Click:
    Call General_Error_Trap
End Sub
Private Sub cmdOk_Click()
On Error GoTo cmdOk_Click
Dim retval
If IsNull(Me![txtLogin]) Or IsNull(Me![txtPwd]) Then
    retval = MsgBox("Sorry but the system cannot continue without both a login name and a password. Do you want to try again?", vbCritical + vbYesNo, "Login required")
    If retval = vbYes Then 'try again
        DoCmd.GoToControl "txtLogin"
        Exit Sub
    Else 'no, don't try again so quit system
        retval = MsgBox("The system cannot continue without a login name and password." & Chr(13) & Chr(13) & "Are you sure you want to quit the system?", vbCritical + vbYesNo, "Confirm System Closure")
        If retval = vbYes Then
            MsgBox "The system will now quit", vbCritical + vbOKOnly, "Invalid Login"
            DoCmd.Quit acQuitSaveAll
        Else 'no I don't want to quit system, ie: try again
            DoCmd.GoToControl "txtLogin"
            Exit Sub
        End If
    End If
Else
    DoCmd.Hourglass True
    If LogUserIn(Me![txtLogin], Me![txtPwd]) = True Then
        DoCmd.Close acForm, "FRM_Login" 'shut form as modal
    Else
    End If
    DoCmd.Hourglass False
End If
Exit Sub
cmdOk_Click:
    Call General_Error_Trap
    DoCmd.Hourglass False
    DoCmd.Close acForm, "Excavation_Login" 'this may be better as a simply quit the system, will see, however must shut form as modal
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmdOk_Click
End If
End Sub
Private Sub txtPwd_KeyPress(KeyAscii As Integer)
End Sub
Private Sub txtPwd_LostFocus()
End Sub
