Attribute VB_Name = "modIO"
'****THIS CODE IS FREE TO USE OR MODIFY IN ANY WAY YOU WANT****
'**PLEASE CREDIT MY NAME IF YOU RE-DISTRIBUTE IT IN ANY FORM.**
'** ENJOY! -Jacob Dinardi   jdinardi@excite.com

Dim iBuff As String, jBuff As String
Dim i As Integer, j As Integer, k As Integer


'----------------------------------------
'Name: PutFile
'----------------------------------------
Public Sub PutFile(Container As String, FileName As String, Optional NewFile As String)
'Container WILL BE THE FILE TO PUT THE OTHER FILE (FileName) INSIDE OF
'IF NewFile IS SPECIFIED THE RESULTING FILE IS SAVED TO ANOTHER NAME,
'OTHERWISE THE CONTAINER FILE IS ACTUALLY MODIFIED
i = FreeFile
Open Container For Binary As i
'CREATE A BUFFER TO COPY THE CONTENTS OF THE CONTAINER TO
iBuff = String(LOF(i), Chr(0))
Get i, , iBuff
'NOW... GET THE CONTENTS OF FileName
j = FreeFile
Open FileName For Binary As j
jBuff = String(LOF(j), Chr(0))
Get j, , jBuff
Close j
'NOW... COMBINE THE FILES
iBuff = iBuff & jBuff
If NewFile <> "" Then
  'MAKE THE FILE SPECIFIED BY THE NewFile ARGUMENT
  k = FreeFile
  Open NewFile For Binary As k
  Put k, , iBuff
  Close k
Else
  'NOT GOING TO MAKE A NEW FILE... ADD DATA TO THE CONTAINER FILE
  Put i, , iBuff
  Close i
End If

End Sub


'----------------------------------------
'Name: GetFile
'----------------------------------------
Public Sub GetFile(Combined As String, Size As Long, Optional NewFile As String)
'YOU MUST KNOW THE SIZE OF THE ORIGINAL CONTAINER FILE, SO YOU KNOW WHERE TO START
'READING BITS FROM THE COMBINED FILE. DOING IT THIS WAY MAKES IT SIMPLE TO INCLUDE
'AN INI FILE FOR EXAMPLE, BECAUSE IT'S SIZE IS FREE TO CHANGE AND IT IS STILL EASILY
'LOCATED. THIS EXAMPLE IS MEANT TO BE SIMPLE AND TO THE POINT, IF YOU WANT TO COMBINE
'MULTIPLE FILES (MORE THAN 2) YOU'LL HAVE TO CREATE MARKER BITS IN THE FILE ALSO TO
'ALLOW YOURSELF TO LOCATE START/STOP POSITIONS, AND OF COURSE ADD CODE TO HANDLE IT.
i = FreeFile
Open Combined For Binary As i
iBuff = String(LOF(i), Chr(0))
Get i, , iBuff
'LENGTH OF COMBINED FILE - Size = LENGTH OF EMBEDDED FILE
iBuff = Trim(Right(iBuff, LOF(i) - Size))
iBuff = Trim(iBuff)
If NewFile <> "" Then
  'MAKE A NEW FILE, OTHERWISE THE DATA IS IN THE BUFFER
  j = FreeFile
  Open NewFile For Binary As j
  Put j, , iBuff
  Close j
End If
Exit Sub
End Sub
