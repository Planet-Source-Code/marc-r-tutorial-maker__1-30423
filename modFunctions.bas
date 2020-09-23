Attribute VB_Name = "modFunctions"
Public Function GetFileNameFromPath(strPath) As String
  Dim intX As Integer
  Dim intPlace As Integer
  Dim intLastPlace As Integer
    
  intLastPlace = 0

  For intX = 1 To Len(strPath)
    intPlace = InStr(intLastPlace + 1, strPath, "\")
    
    If intPlace = 0 Then
      GetFileNameFromPath = Right(strPath, Len(strPath) - intLastPlace)
      Exit Function
    Else
      intLastPlace = intPlace
    End If
  Next intX
End Function

Public Function FileExist(ByVal FileName As String) As Boolean
'Determines if a file exists
On Error Resume Next
If Dir(FileName, vbSystem + vbHidden) = "" Then
    FileExist = False
Else
    FileExist = True
End If
End Function
