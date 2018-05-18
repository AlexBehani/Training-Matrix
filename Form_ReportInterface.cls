VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ReportInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CancelCommand_Click()
DoCmd.Close acForm, "ReportInterface"
DoCmd.OpenForm "MainMenu"
End Sub

Private Sub ClosedTraining_Click()
On Error GoTo Err

DoCmd.OutputTo acOutputQuery, "ClosedTraining", acFormatXLSX

Exit Sub
Err:
If Err.Number = 2501 Then Resume Next

End Sub

Private Sub OpenTraining_Click()
On Error GoTo Err

DoCmd.OutputTo acOutputQuery, "OpenTraining", acFormatPDF

Exit Sub
Err:
If Err.Number = 2501 Then Resume Next
End Sub

Private Sub OverdueTraining_Click()
On Error GoTo Err

DoCmd.OutputTo acOutputQuery, "OverdueTraining", acFormatXLSX

Exit Sub
Err:
If Err.Number = 2501 Then Resume Next
End Sub
