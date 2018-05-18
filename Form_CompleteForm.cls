VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CompleteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub EmployeeCombo_AfterUpdate()
Dim row As String
    
row = "SELECT Employee, SOP, Revision, DueDate, Completed, Group, id" & _
" FROM TrainingTable WHERE Inactive= False AND Employee = '" & Me.EmployeeCombo & "'" & _
"ORDER BY Employee, DueDate DESC;"

Me.TrainingList.RowSource = row
Me.TrainingList.Requery

Me.GroupCombo = ""
Me.SOPCombo = ""
    
    
End Sub

Private Sub ExitCommand_Click()
DoCmd.Close acForm, "CompleteForm"
DoCmd.OpenForm "MainMenu"
End Sub

Private Sub Form_Load()
Dim row As String

row = "SELECT Employee, SOP, Revision, DueDate, Completed, Group, id" & _
" FROM TrainingTable WHERE Inactive= False " & _
" ORDER BY Employee, DueDate DESC;"
'MsgBox row

Me.TrainingList.RowSource = row

End Sub

Private Sub GroupCombo_AfterUpdate()
Dim row As String

row = "SELECT Employee, SOP, Revision, DueDate, Completed, Group, id" & _
" FROM TrainingTable WHERE (((TrainingTable.Group) = '" & Me.GroupCombo & "') And ((TrainingTable.Inactive) = False)) " & _
"ORDER BY TrainingTable.DueDate DESC;"

Me.TrainingList.RowSource = row
Me.TrainingList.Requery

Me.SOPCombo = ""
Me.EmployeeCombo = ""
  End Sub

Private Sub RemoveFilterCommand_Click()
Dim row As String

row = "SELECT Employee, SOP, Revision, DueDate, Completed, Group, id" & _
" FROM TrainingTable WHERE Inactive= False " & _
" ORDER BY Employee, DueDate DESC;"


Me.TrainingList.RowSource = row
Me.TrainingList.Requery
Me.SOPCombo = ""
Me.GroupCombo = ""
Me.EmployeeCombo = ""

End Sub

Private Sub SOPCombo_AfterUpdate()
Dim row As String



row = "SELECT Employee, SOP, Revision, DueDate, Completed, Group, id" & _
" FROM TrainingTable WHERE (((TrainingTable.SOP) = '" & Me.SOPCombo & "') And ((TrainingTable.Inactive) = False)) " & _
"ORDER BY TrainingTable.DueDate DESC;"

Me.TrainingList.RowSource = row
Me.TrainingList.Requery
Me.EmployeeCombo = ""
Me.GroupCombo = ""
End Sub

Private Sub TrainingList_DblClick(Cancel As Integer)
Set Trn = New TrainingClass
Trn.LoadID (Me.TrainingList.Column(6))

If Me.TrainingList.Column(4) = True Then
    Trn.Complete
    Set Trn = Nothing
    DoCmd.Beep
    Me.TrainingList.Requery

Else
    DoCmd.OpenForm "DateCompleted"

End If







End Sub
