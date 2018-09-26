VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmailClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public EmailAdd As String
Public GracePeriod As Integer



Public Sub EmployeeTrigger(rs As Recordset, Fname As String, FullName As String, _
                            EmailAddress As String)
                            
 Dim db As Database
 Dim email As Recordset
 Dim EmailCont As String
 EmailCont = "Hi " & Fname & vbNewLine & vbNewLine & _
 "You are assigned to be trained on the following SOP(s)" & _
vbNewLine & vbNewLine
 
If (rs.RecordCount > 0) Then
rs.MoveFirst
Do While Not rs.EOF

    EmailCont = EmailCont & "SOP: " & rs!SOPName & _
    ", Tiltle: " & rs!SOPTitle & ", Revision: " & rs!CurrentRev & vbNewLine & vbNewLine
    rs.MoveNext
Loop

EmailCont = EmailCont & vbNewLine & vbNewLine & _
"The due date for completing the course(s) is " & _
DateAdd("d", GracePeriod, Date) & _
vbNewLine & vbNewLine & "Best regards"

Set db = CurrentDb
Set email = db.OpenRecordset("EmailTable")

email.AddNew
email!EmailAddress = EmailAddress
email!EmailSubject = "Training"
email!EmailContent = EmailCont
email!status = "Pending"
email!FullName = FullName
email!EmailTopic = "New training"
email.Update


End If


Set db = Nothing
Set email = Nothing
Set rs = Nothing
                            
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


Public Sub RevisionTrigger(rs As Recordset, rev As String)

Dim db As Database
Dim email As Recordset
Dim EmailAdd As String
If (rs.RecordCount > 0) Then
    Set db = CurrentDb
    Set email = db.OpenRecordset("EmailTable")
    rs.MoveFirst
    Do While Not rs.EOF
        EmailAdd = ExtractEmailAdd(rs!FullName)
        If Nz(EmailAdd, "") <> "" Then
        
            email.AddNew
            email!EmailAddress = EmailAdd
            email!EmailSubject = rs!SOPName & " " & rs!SOPTitle
            email!EmailContent = "Hi " & rs!FullName & vbNewLine & vbNewLine & _
                                "Please be informed that you have been assinged training on the following course" & _
                                vbNewLine & vbNewLine & "* " & rs!SOPName & " revision : " & rev & _
                                vbNewLine & vbNewLine & "The due date is " & DateAdd("d", GracePeriod, Date) & _
                                vbNewLine & vbNewLine & "Best regards"
            email!status = "Pending"
            email!FullName = rs!FullName
            email!EmailTopic = "Training - " & rs!SOPName
            email.Update
    
        End If
        
        rs.MoveNext
    Loop
    
End If
Set db = Nothing
Set email = Nothing
Set rs = Nothing

End Sub


Function ExtractEmailAdd(FullName As String) As String

Dim db As Database
Dim rs As Recordset

Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT EmailAdd FROM Employee WHERE FullName='" & FullName & "'")

If rs.RecordCount > 0 Then
rs.MoveFirst
ExtractEmailAdd = Nz(rs!EmailAdd, "")
Else
ExtractEmailAdd = ""
End If

Set db = Nothing
Set rs = Nothing
End Function


