Attribute VB_Name = "LookUpEmployeeModule"
Option Compare Database
Option Explicit

Sub AddEmployee(FullName As String)

Dim db As Database
Dim LEmployee As Recordset

Set db = CurrentDb
Set LEmployee = db.OpenRecordset("LookupEmployees")

LEmployee.AddNew
LEmployee!FullName = FullName
LEmployee!Inactive = False
LEmployee.Update


Set LEmployee = Nothing
Set db = Nothing


End Sub


Sub UpdateFullName(FullName As String, id As Integer)

Dim db As Database
Dim LEmployee As Recordset

Set db = CurrentDb
Set LEmployee = db.OpenRecordset("SELECT FullName FROM LookupEmployees WHERE ID=" & id)

If LEmployee.RecordCount > 0 Then
    LEmployee.MoveFirst
    LEmployee.Edit
    LEmployee!FullName = FullName
    LEmployee.Update
End If

Set db = Nothing
Set LEmployee = Nothing


End Sub

Sub ArchiveEmployee(id As Integer)

Dim db As Database
Dim LEmployee As Recordset

Set db = CurrentDb
Set LEmployee = db.OpenRecordset("SELECT Inactive FROM LookupEmployees WHERE ID=" & id)

If LEmployee.RecordCount > 0 Then
    LEmployee.MoveFirst
    LEmployee.Edit
    LEmployee!Inactive = Not (LEmployee!Inactive)
    LEmployee.Update
End If

Set db = Nothing
Set LEmployee = Nothing
End Sub
