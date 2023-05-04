Sub main
	Dim matchingPattern As String = InputBox("Enter the matching pattern:")
	Dim newPattern As String = InputBox("Enter the new pattern:")
	Dim warning = MsgBox("Warning! are you sure you want to replace '"+ matchingPattern + "' with '" + newPattern + "'?", vbOKCancel, " warning ")
	
	If warning = vbCancel Then Exit Sub
	
	'Get the current directory
    Dim folderPath As String = ThisDoc.Path
	Dim invApp As Inventor.Application = ThisApplication

    ' Get the list of files in the folder
    Dim files As String() = System.IO.Directory.GetFiles(folderPath, matchingPattern)
    
    ' Loop through the files and rename them
    For Each file As String In files
        Dim newName As String = Replace(file, matchingPattern, newPattern, , , CompareMethod.Text)
        System.IO.File.Move(fileName, newName)
    Next
    
    ' Refresh the file list
    invApp.FileManager.RefreshAllDocuments()
End Sub
