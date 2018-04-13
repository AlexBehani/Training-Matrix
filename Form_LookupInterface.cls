VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LookupInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command2_Click()
DoCmd.Close acForm, "LookupInterface"
DoCmd.OpenForm "EmployeeForm"
End Sub

Private Sub Command3_Click()

    DoCmd.Close acForm, "LookupInterface"
    DoCmd.OpenForm "MainMenu"
End Sub

Private Sub EmployeeCommand_Click()
DoCmd.Close acForm, "LookupInterface"
DoCmd.OpenForm "EmployeeForm"
End Sub
