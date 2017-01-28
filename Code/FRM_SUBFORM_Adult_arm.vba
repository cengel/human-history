Option Compare Database
Option Explicit
Private Sub cmdCompleteLHand_Click()
On Error GoTo err_completeLHand
    Me!Metacarpal_1_left = True
    Me!Metacarpal_2_left = True
    Me!Metacarpal_3_left = True
    Me!Metacarpal_4_left = True
    Me!Metacarpal_5_left = True
    Me!Proximal_phalanx_1_left = True
    Me!Distal_phalanx_1_left = True
    Me![Proximal_phalanges_2-5_left] = 4
    Me![Middle_phalanges_2-5_left] = 4
    Me![Distal_phalanges_2-5_left] = 4
Exit Sub
err_completeLHand:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdCompleteRHand_Click()
On Error GoTo err_completeRHand
    Me!Metacarpal_1_right = True
    Me!Metacarpal_2_right = True
    Me!Metacarpal_3_right = True
    Me!Metacarpal_4_right = True
    Me!Metacarpal_5_right = True
    Me!Proximal_phalanx_1_right = True
    Me!Distal_phalanx_1_right = True
    Me![Proximal_phalanges_2-5_right] = 4
    Me![Middle_phalanges_2-5_right] = 4
    Me![Distal_phalanges_2-5_right] = 4
Exit Sub
err_completeRHand:
    Call General_Error_Trap
    Exit Sub
End Sub
