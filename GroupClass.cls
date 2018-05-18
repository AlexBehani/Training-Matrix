VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GroupClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public GroupName As String
Public id As Long

Sub AddLGroup(GroupName As String)

Dim db As Database
Dim LGroup As Recordset

Set db = CurrentDb
Set LGroup = db.OpenRecordset("Groups")

LGroup.AddNew
LGroup!GroupName = GroupName
'LGroup!Inactive = False
LGroup.Update


Set LGroup = Nothing
Set db = Nothing

End Sub

Function UniqueGroup(Grp As String) As Boolean

Dim db As Database
Dim Group As Recordset

Set db = CurrentDb

Set Group = db.OpenRecordset("SELECT id, GroupName FROM Groups WHERE GroupName = '" & Grp & "'")

If Group.RecordCount > 0 Then
    UniqueGroup = False
Else
    UniqueGroup = True
End If

Set db = Nothing
Set Group = Nothing

End Function

