Dim ModulePath As String
Dim thisTarget As Workbook
Dim thisName As String
Dim FileVersion(4) As String
Dim FileDownloadReq(4) As Boolean
Dim FileDownloadLink(4) As String
Dim FileName(4) As String
Dim newHour As String
Dim newMinute As String
Dim newSecond As String

Function SetCurrentVersionNumber()

FileVersion(1) = "0.1a"
FileVersion(2) = "0.1"
FileVersion(3) = "0.1"

FileDownloadReq(1) = False
FileDownloadReq(2) = False
FileDownloadReq(3) = False

FileDownloadLink(1) = "https://drive.google.com/uc?id=1qmjvsaQeoLPHDH-8a_bzXuC3WKQ_bQoZ&authuser=0&export=download"
FileDownloadLink(2) = "https://drive.google.com/uc?id=1qZsI8KvTD2EtD-yG99-u_xpYrBdIkbC_&authuser=0&export=download"
FileDownloadLink(3) = "https://drive.google.com/uc?id=1AmvEu8fInbab0AexRlAuRRoA1X8N-Yr-&authuser=0&export=download"

FileName(1) = "\FileImport.txt"
FileName(2) = "\RequireAssets.txt"
FileName(3) = "\FileVersionManager.txt"

End Function

Function GetFileUpdates()

If (FileVersionManager.GetVersionNumber(1) <> FileVersion(1)) Then
FileDownloadReq(1) = True
End If

If (FileVersionManager.GetVersionNumber(2) <> FileVersion(2)) Then
FileDownloadReq(2) = True
End If

If (FileVersionManager.GetVersionNumber(3) <> FileVersion(3)) Then
FileDownloadReq(3) = True
End If

UpdateFiles

End Function

Function UpdateFiles()

Set thisTarget = ActiveWorkbook
thisName = thisTarget.Name

For I = 1 To 3
If (FileDownloadReq(I) = True) Then
Shell ("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe -url" & FileDownloadLink(I))
ModulePath = FileImport.ChromeDownloadFolder() & FileName(I)

newHour = Hour(Now())
newMinute = Minute(Now())
newSecond = Second(Now()) + 2
waitTime = TimeSerial(newHour, newMinute, newSecond)
Application.Wait waitTime

thisTarget.VBProject.VBComponents.Import ModulePath
FileImport.DeleteFile (FileName(I))
End If
Next I

End Function
