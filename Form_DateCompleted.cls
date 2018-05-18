VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DateCompleted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub OKCommand_Click()
Trn.Complete (Me.DateCompleted)

Set Trn = Nothing
DoCmd.Beep
Forms!CompleteForm!TrainingList.Requery
DoCmd.Close acForm, "DateCompleted"

End Sub
