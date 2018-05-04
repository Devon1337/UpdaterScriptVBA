Dim VersionID(4) As String

Function Setup()
Call SetVersionNumber(3, "0.1")
End Function

Function GetVersionNumber(Index As Integer) As String
GetVersionNumber = VersionID(Index)
End Function

Function SetVersionNumber(Index As Integer, VersionID As String)
VersionID(Index) = VersionID
End Function

Function UpdateOpen(VersionID As String, UpdateVersionID As String) As Boolean

If (VersionID = UpdateVersionID) Then

Else

End If


End Function
