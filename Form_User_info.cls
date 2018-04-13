VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_User_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command19_Click()

End Sub

Private Sub Command6_Click()
On Error GoTo handle_err
DoCmd.Beep
If (MsgBox("You are about to reset " & SUser.GetFullName & "'s Password!" & vbCrLf & vbCrLf & "Are you sure?", vbYesNo, "Warning")) = vbNo Then
Exit Sub
Else:
Call Reset_password(SUser.UserID)
MsgBox SUser.GetFullName & "'s password has changed to welcome", vbDefaultButton1, "Password's Changed"
End If
If (MsgBox("Are you through with the form?", vbYesNo, "Finish?")) = vbYes Then
DoCmd.Close acForm, "User_info"
DoCmd.OpenForm "User_Administration"
End If
'Set SUser = Nothing
Exit Sub
handle_err:
MsgBox Err.Description
End Sub



Private Sub Command8_Click()
DoCmd.Close acForm, "User_info"
DoCmd.OpenForm "User_Administration"
End Sub

Private Sub Delete_command_Click()
On Error GoTo handle_err


If (MsgBox("You are about to delete " & SUser.GetFullName & " from login list!" & vbCrLf & vbCrLf & _
"Are you sure?", vbYesNo, "Warning")) = vbNo Then
Exit Sub
Else:
'MsgBox Fname & " " & Lname & "Acees level = " & Accss_level
Call DeleteUser(SUser.UserID)
   
MsgBox SUser.GetFullName & " has been removed from Login list"
DoCmd.Close acForm, "User_info"
DoCmd.OpenForm "User_Administration"
End If
Exit Sub
handle_err:
MsgBox Err.Description
End Sub

Private Sub ExitCommand_Click()
DoCmd.Close acForm, "User_info"
DoCmd.OpenForm "User_Administration"
End Sub

Private Sub Form_Close()
If Not (SUser Is Nothing) Then Set SUser = Nothing
End Sub

Private Sub Form_Load()
On Error GoTo Err
'Dim FrmAccss As Integer
'FrmAccss = 1
'Call Access_Check(FrmAccss, Form.Name)
Me.Fname = SUser.Fname
Me.Lname = SUser.Lname
Exit Sub
Err:
Resume Next
End Sub

Private Sub SaveButton_Click()
'On Error GoTo handle_err
'Dim FName As String
'Dim LName As String
'Dim Accss_level As Integer
If Nz(Me.Fname, "") = "" Then

    MsgBox "The First Name box is empty!", vbCritical, "First Name"
    Exit Sub
End If
If Nz(Me.Lname, "") = "" Then

    MsgBox "The Last Name box is empty!", vbCritical, "Last Name"
    Exit Sub
End If

'If Nz(Me.A_level, "") = "" Then
'    MsgBox "The Access Level combo is empty!", vbCritical, "Last Name"
'    Exit Sub
'End If
'FName = Me.FName
'LName = Me.LName
'Accss_level = Me.A_level

    If (MsgBox("You are about to change " & SUser.GetFullName & "'s information!" & vbCrLf & vbCrLf & "Are you sure?", vbYesNo, "Warning")) = vbNo Then
    Exit Sub
    Else:

    Call Change_User_info(Me.Fname, Me.Lname, SUser.UserID)
    
    End If
'MsgBox "It has changed to:" & vbCrLf & vbCrLf & UFnme & " " & _
'Ulnme & vbCrLf & "Access Level = " & UACCSS, vbDefaultButton1, "Update"
'End If
If (MsgBox("Are you through with the form?", vbYesNo, "Finish?")) = vbYes Then
DoCmd.Close acForm, "User_info"
DoCmd.OpenForm "User_Administration"
End If
'Set SUser = Nothing
Exit Sub

handle_err:
MsgBox Err.Description
'Set SUser = Nothing
End Sub
