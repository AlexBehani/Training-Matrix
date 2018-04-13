Attribute VB_Name = "Errors"
Option Compare Database

Public Sub Errs(Description As String, Lastdll As String, _
                    ErrorNumber As Integer, Source As String)
                    
Dim db As Database
Dim Errors As Recordset

Set db = CurrentDb
Set Errors = db.OpenRecordset("ErrorCaptured")

Errors.AddNew
Errors!Description = Description
Errors!ErrorlastDllError = Lastdll
Errors!ErrorNumber = ErrorNumber
Errors!Source = Source
Errors!DateTime = Now
If Not (CUser Is Nothing) Then Errors!User = CUser.FullName
'Errors!User = Nz(CUser.FullName, "")
Errors.Update

Set Errors = Nothing
Set db = Nothing
                    
                    
                    
End Sub

