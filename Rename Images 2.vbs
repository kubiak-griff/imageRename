Dim objFSO, objFolder, objFile, i, j
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\ImagesToBeRenamed")

i = 368
j = 1

Dim warning1

warning1 = "Please ensure that the following point below are met: 	1.) No original images are in C:\ImagesToBeRenamed as changes cannot be undone.           	2.) The nnumber of the images in this folder are divisble by the number of sides per sample.		3.) All images are reviewed so that they are aligned with the number of samples (i.e. no duplicates,missing sides etc.)	Click OK to Proceed		Click Cancel to stop this and complete the above before running this again"

Dim response
response = MsgBox(warning1, vbExclamation + vbOKCancel, "Warning")
 
If response = vbOK Then
    MsgBox "Program will now proceed.", vbInformation, "Notice"
Else
    MsgBox "Program will now exit.", vbInformation, "Notice"
    WScript.Quit
End If

For Each objFile In objFolder.Files
    If InStr(objFile.Name, ".jpg") > 0 Then
        If j MOD 6 = 1 Then
            i = i + 1
        End If
	objFile.Name = "DUT" & i & "_ "& j & ".jpg"
        j = j + 1
	If j MOD 6 = 1 Then
	j = 1
	End If	

    End If
Next


