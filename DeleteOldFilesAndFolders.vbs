Option Explicit

' Dim sParameters
' sParameters = "C:\Users\xxx\AppData\Local\Temp\;2;True;True;C:\Users\DARM\Downloads\"
' DeleteOldFilesAndFolders(sParameters) 

Function DeleteOldFilesAndFolders (sParameters)
  Dim path, days, deleteFiles, deleteFolders, arrParameters, result, filePath, folderPath, folderLogPath
  Dim fso, folder, files, file, subfolders, subfolder
  Dim dateLastModified
  Dim logFile, logFileNamePath 'Variables para el archivo de log
  Dim dateProcess
  arrParameters = split(sParameters,";")
  'Get the arguments from the command line
  path = arrParameters(0) 'The first argument is the path
  days = arrParameters(1) 'The second argument is the number of days
  deleteFiles = arrParameters(2) 'The third argument is a boolean value for deleting files
  deleteFolders = arrParameters(3) 'The fourth argument is a boolean value for deleting folders
  folderLogPath = arrParameters(4)
  
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set folder = fso.GetFolder(path)
  Set files = folder.Files
  Set subfolders = folder.SubFolders

  dateProcess = ISODateTime(Now)
    
  'Create a log file with the current date and time in the name
  logFileNamePath = folderLogPath & "DeleteOldFilesAndFolders_" & Replace(ISODate(Now),":",";") & ".log"

  If fso.FileExists(logFileNamePath) Then
    Set logFile = fso.OpenTextFile(logFileNamePath, 8)
  Else
    Set logFile = fso.CreateTextFile(logFileNamePath, True)
  End If
  
  
  If deleteFiles Then
    For Each file in files
      dateLastModified = file.DateLastModified
      filePath = file.Path
      If DateDiff("d", dateLastModified, Date) >= CInt(days) Then
        On Error Resume Next 'Ignore errors and continue execution
        
        file.Delete True
        If Err.Number <> 0 Then 'Check if there was an error
          logFile.WriteLine dateProcess & "ERROR;Deleting file:  " & filePath & ";" & Err.Description 'Write the error message to the log file
          Err.Clear 'Clear the error object
        Else
          logFile.WriteLine dateProcess & ";INFO;Deleting file:  " & filePath 'Write the message to the log file
        End If
        On Error GoTo 0 'Restore default error handling
      Else
        'logFile.WriteLine dateProcess & ";The File " & filePath & ": have " & DateDiff("d", dateLastModified, Date) & " Days old" 'Write the message to the log file
      End If
    Next
  End If
  
  

  If deleteFolders Then
    For Each subfolder in subfolders
      folderPath = subfolder.Path
      dateLastModified = subfolder.DateLastModified
      If DateDiff("d", dateLastModified, Date) >= CInt(days) Then
        On Error Resume Next 'Ignore errors and continue execution
        subfolder.Delete True
        If Err.Number <> 0 Then 'Check if there was an error
          logFile.WriteLine dateProcess & ";ERROR;Deleting folder:" & folderPath & ";" & Err.Description 'Write the error message to the log file
          Err.Clear 'Clear the error object
        Else
          logFile.WriteLine dateProcess & ";INFO;Deleting folder:" & folderPath 'Write the message to the log file 
        End If
        On Error GoTo 0 'Restore default error handling
      Else
        ' logFile.WriteLine dateProcess & ";The folder " & folderPath  & ": have " & DateDiff("d", dateLastModified, Date) & " Days old" 'Write the message to the log file
      End If
    Next
  End If

  Set fso = Nothing  
  logFile.Close 'Close the log file
 
  If Err.Number = 0 Then
  	DeleteOldFilesAndFolders = "Ok"
  Else
    DeleteOldFilesAndFolders = Err.Description
  End If
End Function


Function ISODateTime (d)
  ISODateTime = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2) & " " & Right("0" & Hour(d), 2) & ":" & Right("0" & Minute(d), 2) & ":" & Right("0" & Second(d), 2)
End Function

Function ISODate (d)
  ISODate = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2) 
End Function
