Public Function DeleteFile(str As String) As String
' command to a file
' the target folder is in String named "path"
Dim FSO
Dim path As String

' Source File Location
path = str

' Set Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FileExists(path) Then
    FSO.DeleteFile path, True
    DeleteFile = "Success"
Else
    DeleteFile = "Not Found"
End If
End Function
