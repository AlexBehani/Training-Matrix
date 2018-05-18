VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrainingClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public FullName As String
Public GroupName As String
Public SOPName As String
Public SOPTitle As String
Public Revision As String
Public DueDate As Date
Public Supervisor As String
Public Completed As Boolean
Public CompletedDate As Date
Public TrnId As Integer


Sub LoadEmployee(FullN As String, GroupN As String, _
                    Superv As String)

Let FullName = FullN
Let GroupName = GroupN
Let Supervisor = Superv

End Sub

Sub LoadRevision(Rev As String, SOP As String)

Let Revision = Rev
Let SOPName = SOP

End Sub


Sub EmployeeTrigger()

Dim db As Database
Dim Training As Recordset
Dim GSR As Recordset
Dim Str As String

Str = "SELECT Groups.GroupName, SOP.SOPName, SOP.SOPTitle, SOP.CurrentRev " & _
"FROM Revision INNER JOIN (Groups INNER JOIN SOP ON Groups.GroupName = SOP.F_GroupName) ON Revision.SOPName = SOP.SOPName " & _
"WHERE (((SOP.Inactive)=No)) AND GroupName='" & GroupName & "'"

Set db = CurrentDb
'Set Training = db.OpenRecordset(Str)
Set GSR = db.OpenRecordset(Str)

If GSR.RecordCount > 0 Then
    
    GSR.MoveFirst
    Set Training = db.OpenRecordset("TrainingTable")
    Do While Not GSR.EOF
        Training.AddNew
        With Training
            !Employee = FullName
            !Group = GroupName
            !SOP = GSR!SOPName
            !Revision = GSR!CurrentRev
            !DueDate = DateAdd("d", 14, Date)
            !Completed = False
        End With
        Training.Update
        GSR.MoveNext
    Loop
    
    
End If

Set db = Nothing
Set GSR = Nothing
Set Training = Nothing


End Sub


Sub RevisionTrigger()

Dim db As Database
Dim Training As Recordset
Dim GSR As Recordset
Dim Str As String

Str = "SELECT DISTINCT Employee.FullName, Groups.GroupName, SOP.SOPName, SOP.SOPTitle, Employee.Supervisor " & _
"FROM Revision INNER JOIN ((Groups INNER JOIN Employee ON Groups.GroupName = Employee.GroupName) " & _
"INNER JOIN SOP ON Groups.GroupName = SOP.F_GroupName) ON Revision.SOPName = SOP.SOPName " & _
"WHERE (((Employee.Inactive)=False) AND ((SOP.Inactive)=False)) AND SOP.SOPName='" & SOPName & "'"

Set db = CurrentDb
'Set Training = db.OpenRecordset(Str)
Set GSR = db.OpenRecordset(Str)

If GSR.RecordCount > 0 Then
    
    GSR.MoveFirst
    Set Training = db.OpenRecordset("TrainingTable")
    Do While Not GSR.EOF
        Training.AddNew
        With Training
            !Employee = GSR!FullName
            !Group = GSR!GroupName
            !SOP = SOPName
            !Revision = Revision
            !DueDate = DateAdd("d", 14, Date)
            !Completed = False
        End With
        Training.Update
        GSR.MoveNext
    Loop
    
    
End If

Set db = Nothing
Set GSR = Nothing
Set Training = Nothing


End Sub



Sub ArchiveTrainingRecord()

Dim db As Database
Dim Training As Recordset
Dim GSR As Recordset
Dim Str As String

Str = "SELECT id, Inactive " & _
"FROM TrainingTable WHERE" & _
" SOP='" & SOPName & "' AND Revision ='" & Revision & "' AND Completed=False"

Set db = CurrentDb
Set Training = db.OpenRecordset(Str)

If Training.RecordCount > 0 Then
    
    Training.MoveFirst
    Do While Not Training.EOF
        Training.Edit
        With Training
            !Inactive = True
        End With
        Training.Update
        Training.MoveNext
    Loop
    
    
End If

Set db = Nothing
Set GSR = Nothing
Set Training = Nothing


End Sub

Sub Complete(Optional DateCompleted As Date)
Dim db As Database
Dim TT As Recordset

Set db = CurrentDb
Set TT = db.OpenRecordset("SELECT Completed,DateCompleted FROM TrainingTable WHERE id =" & TrnId)
If TT.RecordCount > 0 Then
    If DateCompleted = "12:00:00 AM" Then
        TT.MoveFirst
        TT.Edit
        TT!Completed = Not (TT!Completed)
        TT!DateCompleted = Null
        TT.Update
    Else
        TT.MoveFirst
        TT.Edit
        TT!Completed = Not (TT!Completed)
        TT!DateCompleted = DateCompleted
        TT.Update
    End If

End If
Set db = Nothing
Set TT = Nothing



End Sub


Sub LoadID(id As Integer)
Let TrnId = id
End Sub
