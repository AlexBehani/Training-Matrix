VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EmployeeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub Command6_Click()
Dim FN As String
FN = InputBox("Please enter the Full Name", "Employee name?")
If Nz(FN, "") = "" Then
    MsgBox "Please enter a valid name", vbCritical, "Entry was empty"
    Exit Sub
End If

If (MsgBox(FN & vbNewLine & "Is the name correct?", vbQuestion + vbYesNo, "?") = vbNo) Then
MsgBox "Please reenter the full name", vbInformation, "Try again"
Exit Sub
Else
If Not (UniqueEmployee(FN)) Then
    MsgBox FN & " is already entered", vbCritical, ""
    Exit Sub
End If

Call AddEmployee(FN)
Me.EmployeeList.Requery
End If

End Sub

Private Sub EmployeeList_DblClick(Cancel As Integer)
Dim UpdatedFullName As String


If (MsgBox("If you would like to Archieve/Restore the employee, press yes" & vbNewLine & _
            "in case you would like to edit information, please press no", vbQuestion + vbYesNo, "Please answer the question") = vbYes) Then
            
            Call ArchiveEmployee(Me.EmployeeList.Column(2))
            Me.EmployeeList.Requery
            MsgBox "Done", vbInformation, ""
            
            Else
            
            UpdatedFullName = InputBox("Please enter the full name", "Full name")
            If Nz(UpdatedFullName, "") = "" Then
                MsgBox "Please enter a valid name", vbCritical, "Entry was empty"
                Exit Sub
            End If
            
            Call UpdateFullName(UpdatedFullName, Me.EmployeeList.Column(2))
            Me.EmployeeList.Requery
            MsgBox "Done", vbInformation, ""
End If

End Sub

Private Sub ExitCommand_Click()
DoCmd.Close acForm, "EmployeeForm"
DoCmd.OpenForm "MainMenu"
End Sub
