Dim VersionNumber As String
Dim ModulePath As String
Dim thisTarget As Workbook
Dim thisName As String
Dim newHour As String
Dim newMinute As String
Dim newSecond As String
Dim FSO

Function StartUp()
Call FileVersionManager.SetVersionNumber(2, "0.1")
End Function

Function DeleteCurrentModule()
Set FSO = CreateObject("Scripting.FileSystemObject")
 Set thisTarget = ActiveWorkbook
    thisName = thisTarget.Name
    
    newHour = Hour(Now())
    newMinute = Minute(Now())
    newSecond = Second(Now()) + 1
    waitTime = TimeSerial(newHour, newMinute, newSecond)
    Application.Wait waitTime
    
    thisTarget.VBProject.VBComponents("FileImport").Name = "BadFile"
  thisTarget.VBProject.VBComponents.Remove thisTarget.VBProject.VBComponents("BadFile")
  
    newSecond = Second(Now()) + 3
    waitTime = TimeSerial(newHour, newMinute, newSecond)
    Application.Wait waitTime
  
    FileImport.DeleteFile ("\FileImport.txt")
  
    Module1.CompletedUpdate

End Function
