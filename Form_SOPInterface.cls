VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SOPInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub AddCommand_Click()

    DoCmd.Close acForm, "SOPInterface"
    DoCmd.OpenForm "SOPProfile"
End Sub


Private Sub EditCombo_AfterUpdate()

'    Set EmpEdit = New EmployeeClass
'    EmpEdit.LoadEmployee (Me.EditCombo)
'
'    DoCmd.Close acForm, "EmployeeInterface"
'    DoCmd.OpenForm "EditEmployeeProfile"



End Sub

Private Sub ExitCommand_Click()
    DoCmd.Close acForm, "SOPInterface"
    DoCmd.OpenForm "MainMenu"
End Sub
