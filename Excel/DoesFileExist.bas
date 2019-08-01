Public Function DoesFileExist(str As String) As String
' command to check if file exists
' the target folder is in String named "path"
Dim FSO
Dim path As String

' Source File Location
path = str

' Set Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(path) Then
    DoesFileExist = "Exists"
Else
    DoesFileExist = "Not Found"
End If
End Function
