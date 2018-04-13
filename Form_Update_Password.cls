VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Update_Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command5_Click()

End Sub

Private Sub Form_Load()
On Error GoTo Err

Me.UserLabel.Caption = CUser.User
Exit Sub
Err:
Resume Next
End Sub

Private Sub Update_Command_Click()

If (Nz(Me.Password, "") = "" Or Nz(Me.ConfirmPassword, "") = "") Then

    MsgBox "TextBox is empty", vbCritical, "Fill the box"
    Exit Sub
End If

If (Not (Nz(Me.Password, "") = Nz(Me.ConfirmPassword))) Then

    MsgBox "Password confiramtion does not match Password", vbCritical, ""
    Exit Sub
End If

Call UpdatePassword(Me.Password)

MsgBox "Password's been updated!", vbInformation, ""
DoCmd.Close acForm, "Update_Password"
DoCmd.OpenForm "MainMenu"

End Sub
