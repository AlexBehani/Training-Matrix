VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LookupGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub AddCommand_Click()

Dim Group As String
Group = InputBox("Please enter the Group Name", "Group")
If Nz(Group, "") = "" Then
    MsgBox "Please enter a valid name", vbCritical, "Entry was empty"
    Exit Sub
End If

If (MsgBox(Group & vbNewLine & "Is the correct?", vbQuestion + vbYesNo, "?") = vbNo) Then
MsgBox "Please reenter the Group name", vbInformation, "Try again"
Exit Sub
Else
Call AddLGroup(Group)
Me.GroupList.Requery
End If
End Sub

Private Sub ExitCommand_Click()
DoCmd.Close acForm, "LookupGroup"
DoCmd.OpenForm "MainMenu"

End Sub



Private Sub GroupList_DblClick(Cancel As Integer)
Dim UpdatedGroupName As String


If (MsgBox("If you would like to Archieve/Restore the Group, press yes" & vbNewLine & _
            "in case you would like to edit information, please press no", vbQuestion + vbYesNo, "Please answer the question") = vbYes) Then
            
            Call ArchiveLGroup(Me.GroupList.Column(2))
            Me.GroupList.Requery
            MsgBox "Done", vbInformation, ""
            
            Else
            
            UpdatedGroupName = InputBox("Please enter the Group name", "Group name")
            If Nz(UpdatedGroupName, "") = "" Then
                MsgBox "Please enter a valid name", vbCritical, "Entry was empty"
                Exit Sub
            End If
            
            Call UpdateGroupName(UpdatedGroupName, Me.GroupList.Column(2))
            Me.GroupList.Requery
            MsgBox "Done", vbInformation, ""
End If

End Sub
