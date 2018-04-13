Attribute VB_Name = "LookupGroupModule"
Option Compare Database
Option Explicit

Sub AddLGroup(GroupName As String)

Dim db As Database
Dim LGroup As Recordset

Set db = CurrentDb
Set LGroup = db.OpenRecordset("LookupGroups")

LGroup.AddNew
LGroup!Group_Name = GroupName
LGroup!inactive = False
LGroup.Update


Set LGroup = Nothing
Set db = Nothing


End Sub


Sub UpdateGroupName(GroupName As String, id As Integer)

Dim db As Database
Dim LGroup As Recordset

Set db = CurrentDb
Set LGroup = db.OpenRecordset("SELECT Group_Name FROM LookupGroups WHERE ID=" & id)

If LGroup.RecordCount > 0 Then
    LGroup.MoveFirst
    LGroup.Edit
    LGroup!Group_Name = GroupName
    LGroup.Update
End If

Set db = Nothing
Set LGroup = Nothing


End Sub


Sub ArchiveLGroup(id As Integer)

Dim db As Database
Dim LGroup As Recordset

Set db = CurrentDb
Set LGroup = db.OpenRecordset("SELECT Inactive FROM LookupGroups WHERE ID=" & id)

If LGroup.RecordCount > 0 Then
    LGroup.MoveFirst
    LGroup.Edit
    LGroup!inactive = Not (LGroup!inactive)
    LGroup.Update
End If

Set db = Nothing
Set LGroup = Nothing
End Sub

