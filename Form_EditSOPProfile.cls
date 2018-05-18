VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditSOPProfile"
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
    DoCmd.Close acForm, "EditSOPProfile"
    DoCmd.OpenForm "MainMenu"

End Sub

Private Sub Form_Load()
Me.SOPCombo = SOPEdit.SOPName
Me.SOPTitle = SOPEdit.SOPTitle
Me.CRevision = SOPEdit.CurrentRev
Me.GroupList.RowSource = SOPEdit.RowString
Me.DateReleased = RevEdit.DateReleased

End Sub

Private Sub GroupCombo_AfterUpdate()
    Dim i As Integer
    Dim j As Integer
    i = Me.GroupList.ListCount
    If i > 0 Then
    
        For j = 0 To i - 1
            If Me.GroupCombo = Me.GroupList.ItemData(j) Then
                MsgBox Me.GroupCombo & " already included", vbCritical, ""
                Exit Sub
            End If

        Next j
    End If
        Me.GroupList.AddItem (Me.GroupCombo)
        
End Sub

Private Sub GroupList_DblClick(Cancel As Integer)
   If Me.GroupList.ItemsSelected.Count = 0 Then Exit Sub
    Me.GroupList.RemoveItem (Me.GroupList)

End Sub

Private Sub SaveCommand_Click()
Dim i As Integer
Dim j As Integer
Dim arr() As Variant
    'If Em_Field Then Exit Sub
    If Me.GroupList.ListCount = 0 Then
        MsgBox "Please add groups to the Group list", vbCritical, ""
        Exit Sub
    End If
    
    i = Me.GroupList.ListCount
    ReDim arr(i)
    
    For j = 0 To i - 1
        arr(j) = Me.GroupList.ItemData(j)
    Next j
    
'    Set SOP = New SOPClass
'    Set Rev = New RevisionClass
    
    Call RevEdit.LoadRevision(Me.CRevision, Me.SOPCombo, Me.DateReleased)
    Call RevEdit.UpdateRevision
    Call SOPEdit.LoadForm(Me.SOPCombo, Me.SOPTitle, Me.CRevision, Me.DateReleased, arr())
    Call SOPEdit.DeleteSOP(Me.SOPCombo)
    Call SOPEdit.SaveSOP
    
    Set Trn = New TrainingClass
    Call Trn.LoadRevision(Me.CRevision, Me.SOPCombo)
    Call Trn.ArchiveTrainingRecord
    Call Trn.RevisionTrigger
    
    MsgBox "Done", vbInformation, ""
    
    DoCmd.Close acForm, "EditSOPProfile"
    DoCmd.OpenForm "MainMenu"
    
    
End Sub
