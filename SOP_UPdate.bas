Attribute VB_Name = "SOP_UPdate"
Option Compare Database
Option Explicit

Public Sub clean_SOP_Templates()
Dim str1 As String
Dim str2 As String
Dim db As Database

str1 = "DELETE * FROM SOP_I"
str2 = "DELETE * FROM SOP_ii"
Set db = CurrentDb

db.Execute (str1)
db.Execute (str2)

End Sub


Public Sub Load_SOP_I(Groups As String, sop As String)
Dim index As Integer
Dim rawArray() As String
Dim db As Database
Dim rs As Recordset

Set db = CurrentDb
Set rs = db.OpenRecordset("SOP_I")

rawArray = Split(Groups, ";")
ReDim varArray(LBound(rawArray) To UBound(rawArray))
For index = LBound(rawArray) To UBound(rawArray)
'    varArray(index) = rawArray(index)
rs.AddNew
rs!sop = sop
rs!Group = rawArray(index)
rs.Update

Next index

Set db = Nothing
Set rs = Nothing

End Sub

Public Sub Load_SOP_II(rawArray() As Variant, sop As String)
Dim index As Integer
Dim db As Database
Dim rs As Recordset

Set db = CurrentDb
Set rs = db.OpenRecordset("SOP_II")

For index = LBound(rawArray) To UBound(rawArray)
'    varArray(index) = rawArray(index)
rs.AddNew
If Nz(rawArray(index), "") <> "" Then
rs!sop = sop
rs!Group = rawArray(index)
rs.Update
End If
Next index

Set db = Nothing
Set rs = Nothing
End Sub


