VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EmployeeProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim Emp As EmployeeClass

Private Sub CancelCommand_Click()
If Not (Emp Is Nothing) Then Set Emp = Nothing
DoCmd.Close acForm, "EmployeeProfile"
DoCmd.OpenForm "MainMenu"
End Sub

Private Sub Command34_Click()

Set Trn = New TrainingClass
Call Trn.LoadEmployee(Me.Fname & " " & Me.Lname, Me.GroupCombo, Me.SupervisorCombo)
Call Trn.EmployeeTrigger


End Sub

Private Sub SaveCommand_Click()
Set Emp = New EmployeeClass

If Em_Field Then Exit Sub


Call Emp.FormLoad(Me.Fname, Me.Lname, Me.SupervisorCombo, _
                    Me.GroupCombo, Nz(Me.IsSupervisor, False), Nz(Me.Inactive, False))
                    
Call Emp.SaveEmployee
Set Trn = New TrainingClass
Call Trn.LoadEmployee(Me.Fname & " " & Me.Lname, Me.GroupCombo, Me.SupervisorCombo)
Call Trn.EmployeeTrigger

MsgBox "Done", vbInformation, ""
Set Emp = Nothing
DoCmd.Close acForm, "EmployeeProfile"
DoCmd.OpenForm "MainMenu"

End Sub
