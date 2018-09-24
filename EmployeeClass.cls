VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmployeeClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public FirstName As String
Public LastName As String
Public FullName As String
Public Supervisor As String
Public Group As String
Public IsSupervisor As Boolean
Public Inactive As Boolean
Public EmailAdd As String

Public Sub FormLoad(Fname As String, Lname As String, _
                SupervisorVar As String, GroupVar As String, _
                IsSupervisorVar As Boolean, InactiveVar As Boolean, EmailAddress As String)
                
    Let FirstName = Fname
    LastName = Lname
    FullName = Fname & " " & Lname
    Let Supervisor = Nz(SupervisorVar, "Bryan Lewis")
    Group = GroupVar
    IsSupervisor = IsSupervisorVar
    Inactive = InactiveVar
    EmailAdd = EmailAddress
                
End Sub
                
Public Sub LoadEmployee(FlName As String)

Dim db As Database
Dim em As Recordset

Set db = CurrentDb
Set em = db.OpenRecordset("SELECT * FROM Employee WHERE FullName = '" & FlName & "'")
em.MoveFirst
With em
    FirstName = !FirstName
    LastName = !LastName
    FullName = !FullName
    Inactive = !Inactive
    Group = !GroupName
    Supervisor = Nz(!Supervisor, "")
    IsSupervisor = !IsSupervisor
    EmailAdd = Nz(!EmailAdd, "")
    
End With
'em.Update
Set db = Nothing
Set em = Nothing

End Sub

Public Function SaveEmployee() As Boolean
On Error GoTo Err_Handel
Dim db As Database
Dim em As Recordset

SaveEmployee = True
Set db = CurrentDb
Set em = db.OpenRecordset("Employee")
em.AddNew
With em
    !FirstName = FirstName
    !LastName = LastName
    !FullName = FullName
    !Inactive = Inactive
    !GroupName = Group
    !Supervisor = Supervisor
    !IsSupervisor = IsSupervisor
End With
em.Update

Set db = Nothing
Set em = Nothing

Exit Function
Err_Handel:
If Err.Number = 3022 Then
    MsgBox FirstName & " " & LastName & " already existed in database", vbCritical, ""
Else
    MsgBox Err.Number & vbNewLine & Err.Description, vbCritical, ""
End If
SaveEmployee = False
Set db = Nothing
Set em = Nothing

End Function

Public Sub UpdateEmployee()
Dim db As Database
Dim em As Recordset

Set db = CurrentDb
Set em = db.OpenRecordset("SELECT * FROM Employee WHERE FullName = '" & FullName & "'")
em.Edit
With em
'    !FirstName = FirstName
'    !LastName = LastName
'    !FullName = FullName
    !Inactive = Inactive
    !GroupName = Group
    !Supervisor = Supervisor
    !IsSupervisor = IsSupervisor
    !EmailAdd = EmailAdd
End With
em.Update

Set db = Nothing
Set em = Nothing

End Sub

