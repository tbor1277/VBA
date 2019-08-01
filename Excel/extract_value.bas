'
' This script extracts the value inside a parenthesis "()"
' tbor 2019
'
'
Public Function extract_value(str As String) As String
Dim openPos As Integer
Dim closePos As Integer
Dim midBit As String
    On Error Resume Next
openPos = InStr(str, "(")
    On Error Resume Next
closePos = InStr(str, ")")
    On Error Resume Next
midBit = Mid(str, openPos + 1, closePos - openPos - 1)

If openPos <> 0 And Len(midBit) > 0 Then
extract_value = midBit
Else
extract_value = "N/A"
End If


End Function

Public Sub test_value()
MsgBox extract_value("NUMBER(9)")
End Sub
