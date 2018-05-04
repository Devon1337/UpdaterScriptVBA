Dim VersionNumber As String
Dim ModulePath As String
Dim thisTarget As Workbook
Dim thisName As String
Dim newHour As String
Dim newMinute As String
Dim newSecond As String

Dim FSO

Function StartupVersionType()

Call FileVersionManager.SetVersionNumber(1, "0.1a")
FileVersionManager.Setup
RequireAssets.StartUp

Set FSO = CreateObject("Scripting.FileSystemObject")
AutomateImport

End Function

Sub AutomateImport()
    
    ModulePath = ChromeDownloadFolder() & "\Updater.txt"

    Application.DisplayAlerts = False

    Shell ("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe -url https://drive.google.com/uc?id=1Xf64DJuWggjqfTo1aOKQFVglp_S5jVM1&authuser=0&export=download")

    
    Set thisTarget = ActiveWorkbook
    thisName = thisTarget.Name
    
    newHour = Hour(Now())
    newMinute = Minute(Now())
    newSecond = Second(Now()) + 2
    waitTime = TimeSerial(newHour, newMinute, newSecond)
    Application.Wait waitTime

    thisTarget.VBProject.VBComponents.Import ModulePath
    thisTarget.VBProject.VBComponents("Module1").Name = "Updater"

    ActiveWorkbook.Save
    DeleteFile ("\Updater.txt")
    
    Updater.SetCurrentVersionNumber
    Updater.GetFileUpdates
    
End Sub

Sub DeleteFile(Message As String)
FileToDelete = ChromeDownloadFolder() & Message

   If FSO.FileExists(FileToDelete) Then
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If

End Sub

Function ChromeDownloadFolder()
    Dim sPref As String
    Dim iFile As Long, iStart As Long, iEnd As Long
    Dim sBuffer As String, sSearch As String, sDownloads As String

    sPref = Environ("LOCALAPPDATA") & "\Google\Chrome\User Data\Default\Preferences"
    
    sSearch = """download"":{""default_directory"":"

    iFile = FreeFile
    Open sPref For Input As #iFile
        sBuffer = Input$(LOF(iFile), iFile)
    Close #iFile

    iStart = InStr(1, sBuffer, sSearch, vbTextCompare)
    
    iEnd = InStr(iStart + Len(sSearch), sBuffer, ",", vbTextCompare)

    sDownloads = Mid(sBuffer, iStart + Len(sSearch) + 1, iEnd - iStart - Len(sSearch) - 2)

    ChromeDownloadFolder = Replace(sDownloads, "\\", "\")
End Function

Function CompletedUpdate()
Set thisTarget = ActiveWorkbook
    thisName = thisTarget.Name
    
    newHour = Hour(Now())
    newMinute = Minute(Now())
    newSecond = Second(Now()) + 1
    waitTime = TimeSerial(newHour, newMinute, newSecond)
    Application.Wait waitTime

  
  thisTarget.VBProject.VBComponents("Module1").Name = "FileImport"
thisTarget.VBProject.VBComponents.Remove thisTarget.VBProject.VBComponents("Updater")

End Function
