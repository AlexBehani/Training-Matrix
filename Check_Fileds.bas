Attribute VB_Name = "Check_Fileds"
Option Compare Database

Function Em_Field() As Boolean
On Error GoTo err_handel
Em_Field = False
Dim ctrl As Control
Dim EmpyField As Boolean
EmpyField = False



For Each ctrl In Screen.ActiveForm.Controls

    If Len(ctrl.Tag) <> 0 And ctrl.ControlType <> acCheckBox And ctrl.Visible = True Then

        ctrl.BackColor = RGB(250, 250, 250)
        

    End If
Next

For Each ctrl In Screen.ActiveForm.Controls


        
        If Len(ctrl.Tag) <> 0 And ctrl.ControlType <> acCheckBox And ctrl.Visible = True Then

            If IsNull(ctrl) Then

                    ctrl.BackColor = RGB(250, 100, 100)

                    EmpyField = True
            End If
        End If
Next

If EmpyField Then
    MsgBox "Please fill out the required field(s)", vbCritical + vbExclamation, CurrUser
    Em_Field = True
    Exit Function
End If

Exit Function
err_handel:

If Err.Number = 2455 Then
Resume Next
Else
MsgBox Err.Description & Err.Number
End If
End Function
