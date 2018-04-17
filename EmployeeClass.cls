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

Public Sub FormLoad(Fname As String, Lname As String, _
                SupervisorVar As String, GroupVar As String, _
                IsSupervisorVar As Boolean, InactiveVar As Boolean)
                
    Let FirstName = Fname
    LastName = Lname
    FullName = Fname & " " & Lname
    Let Supervisor = Nz(SupervisorVar, "Bryan Lewis")
    Group = GroupVar
    IsSupervisor = IsSupervisorVar
    Inactive = InactiveVar

                
End Sub
                
Public Sub LoadEmployee(FlName As String)

Dim db As Database
Dim em As Recordset

Set db = CurrentDb
Set em = db.OpenRecordset("SELECT * FROM Employee WHERE FullName = '" & FlName & "'")

With em
    FirstName = !FirstName
    LastName = !LastName
    FullName = !FullName
    Inactive = !Inactive
    Group = !GroupName
    Supervisor = !Supervisor
    IsSupervisor = !IsSupervisor
End With
'em.Update
Set db = Nothing
Set em = Nothing

End Sub

Public Sub SaveEmployee()
Dim db As Database
Dim em As Recordset

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

End Sub

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
End With
em.Update

Set db = Nothing
Set em = Nothing

End Sub

