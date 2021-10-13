Move = True 'Set to False for Copy
SrcDir = ".\Temperature"
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set Source = oFSO.GetFolder(SrcDir)
For Each File in Source.Files
  FileName = Right(File,12)
  YYYY = Mid(FileName,1,4)
  MM = Mid(FileName,5,2)
  YearDir = SrcDir & "\" & YYYY & "\"
  MonthDir = YearDir & MM & "\"
  If Not oFSO.FolderExists(YearDir) Then oFSO.CreateFolder(YearDir)
  If Not oFSO.FolderExists(MonthDir) Then oFSO.CreateFolder(MonthDir)
  If Move Then oFSO.MoveFile(File), MonthDir Else oFSO.CopyFile(File), MonthDir
Next