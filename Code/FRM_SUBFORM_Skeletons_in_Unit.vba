Option Compare Database
Option Explicit
Private Sub cmdView_Click()
On Error GoTo Err_cmdView_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stLinkCriteria = "[UnitNumber] = " & Me![UnitNumber] & " AND [Individual Number] = " & Me![txtIndivid]
    stDocName = "FRM_SkeletonDescription"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "FRM_UnitDescription"
Exit_cmdView_Click:
    Exit Sub
Err_cmdView_Click:
    MsgBox Err.Description
    Resume Exit_cmdView_Click
End Sub
