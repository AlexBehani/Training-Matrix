VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditEmployeeProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub CancelCommand_Click()
If Not (EmpEdit Is Nothing) Then Set EmpEdit = Nothing
    DoCmd.Close acForm, "EditEmployeeProfile"
    DoCmd.OpenForm "MainMenu"
End Sub

Private Sub Form_Load()
Me.Fname = EmpEdit.FirstName
Me.Lname = EmpEdit.LastName
Me.SupervisorCombo = EmpEdit.Supervisor
Me.GroupCombo = EmpEdit.Group
Me.IsSupervisor = EmpEdit.IsSupervisor
Me.Inactive = EmpEdit.Inactive
End Sub

Private Sub SaveCommand_Click()

    Dim Condtion As Boolean

    If Em_Field Then Exit Sub
    
    If EmpEdit.Group <> Me.GroupCombo Then Condtion = True Else Condtion = False
    
    Call EmpEdit.FormLoad(Me.Fname, Me.Lname, Me.SupervisorCombo, _
                        Me.GroupCombo, Nz(Me.IsSupervisor, False), Nz(Me.Inactive, False))
    
    Call EmpEdit.UpdateEmployee

    Set EmpEdit = Nothing
    
    If Condtion Then
        Set Trn = New TrainingClass
        Call Trn.LoadEmployee(Me.Fname & " " & Me.Lname, _
                               Me.GroupCombo, Me.SupervisorCombo)
        Call Trn.EmployeeTrigger
    End If
    MsgBox "Done", vbInformation, ""
    DoCmd.Close acForm, "EditEmployeeProfile"
    DoCmd.OpenForm "MainMenu"

End Sub
