Attribute VB_Name = "Login_Module"
Option Compare Database
Public CUser As CurrentUser
Public SUser As SelectedUser



' login=0 User & Password correct, also no need to update
' login=1 User & Password correct, need to update
' login=2 User & password not correct

Public Function Login(UserName As String, Pass As String) As Integer


Dim db As dao.Database
Dim Rs As dao.Recordset
Dim txt As String
Dim PassTxt As String


Set db = CurrentDb

PassTxt = BASE64SHA1(Pass)
Set Rs = db.OpenRecordset("SELECT * FROM Users WHERE UserName='" & UserName & "' AND Password='" & PassTxt & "'")

If (Rs.RecordCount > 0) Then

        If (Rs!pwdrst = 0) Then
        
            Login = 0
        ElseIf (Rs!pwdrst = -1) Then
            Login = 1
        End If
    Set CUser = New CurrentUser
    CUser.User = Nz(Rs!Fname, "") & " " & Nz(Rs!Lname, "")
    CUser.Fname = Nz(Rs!Fname, "")
    CUser.Lname = Nz(Rs!Lname, "")
    
Else
Login = 2
    
End If

db.Close
Set Rs = Nothing
Set db = Nothing

End Function


Public Sub Register_User(Fname As String, Lname As String, Optional var3 As Integer)
Dim db As Database

Dim PWR As String
Dim Rs As Recordset

PWR = "-1"
Set db = CurrentDb
Set Rs = db.OpenRecordset("SELECT * FROM Users WHERE FName = '" & Fname & "' AND LName = '" & Lname & "'")

If Rs.RecordCount > 0 Then
    
    MsgBox "User info is not unique", vbCritical, "Duplicated information"
    Set Rs = Nothing
    Set db = Nothing
    Exit Sub
End If

Set Rs = db.OpenRecordset("Users")
    
    Rs.AddNew
    Rs!Fname = Fname
    Rs!Lname = Lname
    Rs!Password = BASE64SHA1("welcome")
    Rs!pwdrst = -1
    Rs!UserName = Fname & Lname
    Rs.Update
    
MsgBox "New User is added", vbInformation, "Done"
Set Rs = Nothing
Set db = Nothing

End Sub


Public Function Change_User_info(Fname As String, Lname As String, UserID As Integer)
On Error GoTo Err
Dim db As Database
Dim User As Recordset

Set db = CurrentDb
Set User = db.OpenRecordset("SELECT * FROM Users WHERE UserID = " & UserID)
User.MoveFirst
User.Edit
User!Fname = Fname
User!Lname = Lname
User.Update


Set db = Nothing
Set Rs = Nothing

Exit Function
Err:
MsgBox Err.Number, vbCritical, "Error"
Set db = Nothing
Set Rs = Nothing

End Function


Public Sub Reset_password(UserID As Integer)
Dim db As Database
Dim User As Recordset

Set db = CurrentDb
Set User = db.OpenRecordset("SELECT * FROM Users WHERE UserID = " & UserID)

User.MoveFirst
User.Edit
User!Password = BASE64SHA1("welcome")
User!pwdrst = -1
User.Update

Set db = Nothing
Set User = Nothing
End Sub


Public Sub DeleteUser(UserID As Integer)
Dim db As Database
Dim User As Recordset

Set db = CurrentDb
Set User = db.OpenRecordset("SELECT * FROM Users WHERE UserID = " & UserID)

User.MoveFirst
User.Delete



Set db = Nothing
Set User = Nothing
End Sub



Public Sub UpdatePassword(Pass As String)

Dim db As Database
Dim User As Recordset

Set db = CurrentDb
Set User = db.OpenRecordset("SELECT * FROM Users WHERE FName = '" & CUser.Fname & "' AND LName = '" & CUser.Lname & "'")

User.MoveFirst
User.Edit
User!Password = BASE64SHA1(Pass)
User!pwdrst = False
User.Update

Set db = Nothing
Set User = Nothing




End Sub
