VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditRevProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SOP As SOPClass
Dim Rev As RevisionClass



Private Sub CancelCommand_Click()

    If Not (SOP Is Nothing) Then Set SOP = Nothing
    If Not (Rev Is Nothing) Then Set Rev = Nothing
    DoCmd.Close acForm, "EditRevProfile"
    DoCmd.OpenForm "MainMenu"

End Sub

Private Sub Form_Load()
Me.SOPCombo = SOPEdit.SOPName
Me.CRevision = SOPEdit.CurrentRev
Me.DateReleased = RevEdit.DateReleased

End Sub

Private Sub SaveCommand_Click()
Dim Condtion As Boolean

    If SOPEdit.CurrentRev <> Me.CRevision Then Condtion = True Else Condtion = False
    
    If Condtion And RevEdit.DateReleased = Me.DateReleased Then
        MsgBox "Please update the Date Released", vbCritical, ""
        Exit Sub
    End If
    
    Call RevEdit.LoadRevision(Me.CRevision, Me.SOPCombo, Me.DateReleased)
    Call RevEdit.UpdateRevision
    Call SOPEdit.LoadRevForm(Me.CRevision, Me.SOPCombo)
    Call SOPEdit.UpdateRev(Me.SOPCombo)
    
    If Condtion Then
        Set Trn = New TrainingClass
        Call Trn.LoadRevision(Me.CRevision, Me.SOPCombo)
        Call Trn.RevisionTrigger
    End If
    
    MsgBox "Done", vbInformation, ""
    
    DoCmd.Close acForm, "EditRevProfile"
    DoCmd.OpenForm "MainMenu"
    
    
End Sub
