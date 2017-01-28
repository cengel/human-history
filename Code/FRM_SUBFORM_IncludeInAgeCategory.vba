Option Compare Database
Option Explicit
Private Sub Check21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo err_chkInclude
Dim sql
If Me!Check21 = True Then
    sql = "UPDATE [HR_ageing and sexing] SET [IncludeinAgeSexGrouping] = false WHERE [Unit number] = " & Me!Unit & " AND [Individual Number] = " & Me![IndividualNumber] & ";"
    DoCmd.RunSQL sql
Else
    sql = "UPDATE [HR_ageing and sexing] SET [IncludeinAgeSexGrouping] = true WHERE [Unit number] = " & Me!Unit & " AND [Individual Number] = " & Me![IndividualNumber] & ";"
    DoCmd.RunSQL sql
End If
Me.Requery
Exit Sub
err_chkInclude:
    Call General_Error_Trap
End Sub
Private Sub cmdClose_Click()
On Error GoTo err_close
    DoCmd.Close acForm, Me.Name
Exit Sub
err_close:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_open
If Me.OpenArgs <> "" Then
    Me.RecordSource = "SELECT DISTINCT [HR_Skeleton_RelatedTo_Skeleton].[Unit], [HR_Skeleton_RelatedTo_Skeleton].[IndividualNumber], [HR_ageing and sexing].[IncludeinAgeSexGrouping] FROM HR_Skeleton_RelatedTo_Skeleton LEFT JOIN [HR_ageing and sexing] ON ([HR_Skeleton_RelatedTo_Skeleton].[Unit]=[HR_ageing and sexing].[unit number]) AND ([HR_Skeleton_RelatedTo_Skeleton].[IndividualNumber]=[HR_ageing and sexing].[Individual number]) WHERE " & Me.OpenArgs & ";"
End If
Exit Sub
err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
