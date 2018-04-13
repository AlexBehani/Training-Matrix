Attribute VB_Name = "ShiftKey"
Option Compare Database

Function ap_DisableShift()
'On Error GoTo errDisableShift
'Dim db As dao.Database
'Dim prop As dao.Property
'Const conPropNotFound = 3270
'Set db = CurrentDb()
'db.Properties("AllowByPassKey") = False
'Exit Function
'errDisableShift:
'If Err = conPropNotFound Then
'Set prop = db.CreateProperty("AllowByPassKey", _
'dbBoolean, False)
'db.Properties.Append prop
'Resume Next
'Else
'MsgBox "Function 'ap_DisableShift' did not complete successfully."
'Exit Function
'End If
End Function


Function ap_EnableShift()
On Error GoTo errDisableShift
Dim db As dao.Database
Dim prop As dao.Property
Const conPropNotFound = 3270
Set db = CurrentDb()
db.Properties("AllowByPassKey") = True
Exit Function
errDisableShift:
If Err = conPropNotFound Then
Set prop = db.CreateProperty("AllowByPassKey", _
dbBoolean, True)
db.Properties.Append prop
Resume Next
Else
MsgBox "Function 'ap_DisableShift' did not complete successfully."
Exit Function
End If
End Function

