Attribute VB_Name = "LookupSOPModule"
Option Compare Database
Option Explicit

Sub AddLSOP(SOPName As String)

Dim db As Database
Dim LSOP As Recordset

Set db = CurrentDb
Set LSOP = db.OpenRecordset("LookupSOPs")

LSOP.AddNew
LSOP!SOP_Name = SOPName
LSOP!Inactive = False
LSOP.Update


Set LSOP = Nothing
Set db = Nothing


End Sub


Sub UpdateSOPName(SOPName As String, id As Integer)

Dim db As Database
Dim LSOP As Recordset

Set db = CurrentDb
Set LSOP = db.OpenRecordset("SELECT SOP_Name FROM LookupSOPs WHERE ID=" & id)

If LSOP.RecordCount > 0 Then
    LSOP.MoveFirst
    LSOP.Edit
    LSOP!SOP_Name = SOPName
    LSOP.Update
End If

Set db = Nothing
Set LSOP = Nothing


End Sub


Sub ArchiveLSOP(id As Integer)

Dim db As Database
Dim LSOP As Recordset

Set db = CurrentDb
Set LSOP = db.OpenRecordset("SELECT Inactive FROM LookupSOPs WHERE ID=" & id)

If LSOP.RecordCount > 0 Then
    LSOP.MoveFirst
    LSOP.Edit
    LSOP!Inactive = Not (LSOP!Inactive)
    LSOP.Update
End If

Set db = Nothing
Set LSOP = Nothing
End Sub
