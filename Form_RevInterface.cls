VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_RevInterface"
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


Set SOPEdit = New SOPClass
Set RevEdit = New RevisionClass

SOPEdit.LoadRecord (Me.EditCombo)
RevEdit.LoadRecord (Me.EditCombo)
DoCmd.Close acForm, "RevInterface"
DoCmd.OpenForm "EditRevProfile"






End Sub

Private Sub ExitCommand_Click()
    DoCmd.Close acForm, "RevInterface"
    DoCmd.OpenForm "MainMenu"
End Sub
