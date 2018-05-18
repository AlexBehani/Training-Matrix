VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_GroupProfile"
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
    DoCmd.Close acForm, "GroupProfile"
    DoCmd.OpenForm "MainMenu"

End Sub

Private Sub Form_Load()
Set Grp = New GroupClass

End Sub

Private Sub GroupCombo_AfterUpdate()

If Grp.UniqueGroup(Me.GroupCombo) Then
    Grp.AddLGroup (Me.GroupCombo)
    Me.GroupList.Requery
Else
    MsgBox Me.GroupCombo & " is already included in the Group List", vbCritical, ""
End If

        
End Sub

Private Sub GroupList_DblClick(Cancel As Integer)
'   If Me.GroupList.ItemsSelected.Count = 0 Then Exit Sub
'    Me.GroupList.RemoveItem (Me.GroupList)

End Sub

