Attribute VB_Name = "Sha64"
Option Compare Database

Public Function BASE64SHA1(ByVal sTextToHash As String)
On Error GoTo Err
    Dim asc As Object
    Dim enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    Dim bytes() As Byte
    Const cutoff As Integer = 5

    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA1")

    TextToHash = asc.GetBytes_4(sTextToHash)
    SharedSecretKey = asc.GetBytes_4(sTextToHash)
    enc.Key = SharedSecretKey

    bytes = enc.ComputeHash_2((TextToHash))
    BASE64SHA1 = EncodeBase64(bytes)
    BASE64SHA1 = Left(BASE64SHA1, cutoff)

    Set asc = Nothing
    Set enc = Nothing
Exit Function
Err:
MsgBox Err.Description
Resume Next
End Function

Private Function EncodeBase64(ByRef arrData() As Byte) As String

    Dim objXML As Object
    Dim objNode As Object

    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.Text

    Set objNode = Nothing
    Set objXML = Nothing

End Function
'
'Public Function EncodeBase64(ByVal strPlainText As String) As String
'Dim obHash As Object
'On Error GoTo err_Hash
'Set obHash = CreateObject("CAPICOM.HashedData")
'obHash.Algorithm = 0 'that's SHA1 - see CAPICOM documentation for MDx
'obHash.Hash strPlainText
'Hash = obHash.value
'exit_Hash:
'strPlainText = ""
'Set obHash = Nothing
'Exit Function
'err_Hash:
'MsgBox Err.Number & ": " & Err.Description, vbInformation, "Hash error"
'Hash = ""
'Resume exit_Hash
'End Function
