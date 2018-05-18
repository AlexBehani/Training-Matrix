VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SOPProfile"
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
    DoCmd.Close acForm, "SOPProfile"
    DoCmd.OpenForm "MainMenu"

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

    Set SOP = New SOPClass
    If Not (SOP.UniqueSOP(Me.SOPCombo)) Then
        MsgBox "SOP is already entered", vbCritical, ""
        Exit Sub
    End If
    
    If Em_Field Then Exit Sub
        
    If Me.GroupList.ListCount = 0 Then
        MsgBox "Please add groups to the Group list", vbCritical, ""
        Exit Sub
    End If
    
    i = Me.GroupList.ListCount
    ReDim arr(i)
    
    For j = 0 To i - 1
        arr(j) = Me.GroupList.ItemData(j)
    Next j
    

    Set Rev = New RevisionClass
    
    Call Rev.LoadRevision(Me.CRevision, Me.SOPCombo, Me.DateReleased)
    Call Rev.SaveRevision
    Call SOP.LoadForm(Me.SOPCombo, Me.SOPTitle, Me.CRevision, Me.DateReleased, arr())
    Call SOP.SaveSOP
    
    Set Trn = New TrainingClass
    Call Trn.LoadRevision(Me.CRevision, Me.SOPCombo)
    Call Trn.RevisionTrigger
    
    
    MsgBox "Done", vbInformation, ""
    
    
    Set Rev = Nothing
    Set Trn = Nothing
    Set SOP = Nothing
    DoCmd.Close acForm, "SOPProfile"
    DoCmd.OpenForm "MainMenu"
    
    
End Sub
