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
Public GracePeriod As Integer
Public EmailAdd As String


Sub LoadEmployee(FullN As String, GroupN As String, _
                    Superv As String, EmailAddress As String)

Let FullName = FullN
Let GroupName = GroupN
Let Supervisor = Superv
Let EmailAdd = EmailAddress

End Sub

Sub LoadRevision(rev As String, sop As String)

Let Revision = rev
Let SOPName = sop

End Sub


Function EmployeeTrigger() As Recordset

Dim db As Database
Dim Training As Recordset
Dim GSR As Recordset
Dim str As String

str = "SELECT Groups.GroupName, SOP.SOPName, SOP.SOPTitle, SOP.CurrentRev " & _
"FROM Revision INNER JOIN (Groups INNER JOIN SOP ON Groups.GroupName = SOP.F_GroupName) ON Revision.SOPName = SOP.SOPName " & _
"WHERE (((SOP.Inactive)=No)) AND GroupName='" & GroupName & "'"

Call LoadGracePeriod
Set db = CurrentDb
'Set Training = db.OpenRecordset(Str)
Set GSR = db.OpenRecordset(str)

If GSR.RecordCount > 0 Then
    
    GSR.MoveFirst
    Set Training = db.OpenRecordset("TrainingTable")
    Do While Not GSR.EOF
        Training.AddNew
        With Training
            !Employee = FullName
            !Group = GroupName
            !sop = GSR!SOPName
            !Revision = GSR!CurrentRev
            !DueDate = DateAdd("d", GracePeriod, Date)
            !Completed = False
        End With
        Training.Update
        GSR.MoveNext
    Loop
    
   Set EmployeeTrigger = GSR
End If

Set db = Nothing
Set GSR = Nothing
Set Training = Nothing


End Function


Function RevisionTrigger() As Recordset

Dim db As Database
Dim Training As Recordset
Dim GSR As Recordset
Dim str As String

str = "SELECT DISTINCT Employee.FullName, Groups.GroupName, SOP.SOPName, SOP.SOPTitle, Employee.Supervisor " & _
"FROM Revision INNER JOIN ((Groups INNER JOIN Employee ON Groups.GroupName = Employee.GroupName) " & _
"INNER JOIN SOP ON Groups.GroupName = SOP.F_GroupName) ON Revision.SOPName = SOP.SOPName " & _
"WHERE (((Employee.Inactive)=False) AND ((SOP.Inactive)=False)) AND SOP.SOPName='" & SOPName & "'"

Call LoadGracePeriod
Set db = CurrentDb
'Set Training = db.OpenRecordset(Str)
Set GSR = db.OpenRecordset(str)

If GSR.RecordCount > 0 Then
    
    GSR.MoveFirst
    Set Training = db.OpenRecordset("TrainingTable")
    Do While Not GSR.EOF
        Training.AddNew
        With Training
            !Employee = GSR!FullName
            !Group = GSR!GroupName
            !sop = SOPName
            !Revision = Revision
            !DueDate = DateAdd("d", GracePeriod, Date)
            !Completed = False
        End With
        Training.Update
        GSR.MoveNext
    Loop
        
    
    
End If
Set RevisionTrigger = GSR
Set db = Nothing
Set GSR = Nothing
Set Training = Nothing


End Function

Sub DeleteTrainingRecord()

Dim db As Database
Dim RmvRecords As Recordset
Dim str As String
Dim RmvStr As String
Set db = CurrentDb
'str = "DELETE " & _
'"FROM TrainingTable WHERE" & _
'" SOP='" & SOPName & "' AND Revision ='" & Revision & "' AND Completed=False"
str = "SELECT SOP_I.SOP, SOP_I.Group " & _
"FROM SOP_I LEFT JOIN SOP_II ON SOP_I.Group = SOP_II.Group " & _
"WHERE (((SOP_II.Group) Is Null))"

Set RmvRecords = db.OpenRecordset(str)
If RmvRecords.RecordCount > 0 Then
    RmvRecords.MoveFirst
    Do While Not RmvRecords.EOF
        RmvStr = "DELETE * FROM TrainingTable WHERE Completed =False " & _
        "AND SOP ='" & SOPName & "' AND TrainingTable.Group ='" & RmvRecords!Group & "'"
        db.Execute (RmvStr)
 
    RmvRecords.MoveNext
    Loop


End If




Set RmvRecords = Nothing
'db.Execute str

Set db = Nothing



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


Sub ObseleteTraining()
Dim db As Database
Dim TT As Recordset

Set db = CurrentDb
Set TT = db.OpenRecordset("SELECT Inactive FROM TrainingTable WHERE Employee ='" & FullName & "' AND Completed= False")
If TT.RecordCount > 0 Then
    
    TT.MoveFirst
    Do While Not TT.EOF
        TT.Edit
        TT!Inactive = True
        TT.Update
        TT.MoveNext
    Loop

End If
Set db = Nothing
Set TT = Nothing

End Sub

Sub RestoreTraining()
Dim db As Database
Dim TT As Recordset

Set db = CurrentDb
Set TT = db.OpenRecordset("SELECT Inactive FROM TrainingTable WHERE Employee ='" & FullName & "' AND Completed= False")
If TT.RecordCount > 0 Then
    
    TT.MoveFirst
    Do While Not TT.EOF
        TT.Edit
        TT!Inactive = False
        TT.Update
        TT.MoveNext
    Loop

End If
Set db = Nothing
Set TT = Nothing

End Sub

Sub LoadGracePeriod()
Dim db As Database
Dim GP As Recordset

Set db = CurrentDb
Set GP = db.OpenRecordset("GracePeriod", , dbReadOnly)
GP.MoveFirst
Let GracePeriod = GP!G_Period

Set db = Nothing
Set GP = Nothing
End Sub
