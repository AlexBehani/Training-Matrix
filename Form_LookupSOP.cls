VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LookupSOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub AddCommand_Click()
Dim SOP As String
SOP = InputBox("Please enter the SOP Name", "SOP")
If Nz(SOP, "") = "" Then
    MsgBox "Please enter a valid name", vbCritical, "Entry was empty"
    Exit Sub
End If

If (MsgBox(SOP & vbNewLine & "Is the correct?", vbQuestion + vbYesNo, "?") = vbNo) Then
MsgBox "Please reenter the SOP name", vbInformation, "Try again"
Exit Sub
Else
If Not (UniqueSOP(SOP)) Then
    MsgBox "SOP is already entered", vbCritical, ""
    Exit Sub
End If
Call AddLSOP(SOP)
Me.SOPList.Requery
End If
End Sub

Private Sub ExitCommand_Click()
DoCmd.Close acForm, "LookupSOP"
DoCmd.OpenForm "MainMenu"

End Sub

Private Sub SOPList_DblClick(Cancel As Integer)

Dim UpdatedSOPName As String


If (MsgBox("If you would like to Archieve/Restore the SOP, press yes" & vbNewLine & _
            "in case you would like to edit information, please press no", vbQuestion + vbYesNo, "Please answer the question") = vbYes) Then
            
            Call ArchiveLSOP(Me.SOPList.Column(2))
            Me.SOPList.Requery
            MsgBox "Done", vbInformation, ""
            
            Else
            
            UpdatedSOPName = InputBox("Please enter the SOP name", "SOP name")
            If Nz(UpdatedSOPName, "") = "" Then
                MsgBox "Please enter a valid name", vbCritical, "Entry was empty"
                Exit Sub
            End If
            
            Call UpdateSOPName(UpdatedSOPName, Me.SOPList.Column(2))
            Me.SOPList.Requery
            MsgBox "Done", vbInformation, ""
End If
End Sub
