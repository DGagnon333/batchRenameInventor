Sub main
    ' Get the active document
    Dim activeDoc As Document = ThisApplication.ActiveDocument

    ' Check if the document is a part or assembly file
    If Not TypeOf activeDoc Is PartDocument And Not TypeOf activeDoc Is AssemblyDocument Then
        MsgBox("This script is intended to be used with part and assembly documents only.")
        Exit Sub
    End If
	
	Dim matchingPattern As String = InputBox("Enter the matching pattern:")
	
	If matchingPattern = ""
		MsgBox("matching pattern can't be empty")
		Exit Sub
	End If
	
	Dim newPattern As String = InputBox("Enter the new pattern:")
	Dim warning = MsgBox("Warning! are you sure you want to replace '"+ matchingPattern + "' with '" + newPattern + "'?", vbOKCancel, " warning ")
	
	If warning = vbCancel Then 
		Exit Sub
	End If

    ' Get the full path and filename of the active document
    Dim filePath As String = activeDoc.FullFileName
    Dim folderPath As String = System.IO.Path.GetDirectoryName(filePath)

    ' Get the list of files in the folder
    Dim files As String() = System.IO.Directory.GetFiles(folderPath, matchingPattern)

    ' Loop through the files and rename them
    For Each file As String In files
        Dim oldName As String = System.IO.Path.GetFileNameWithoutExtension(File)
        Dim newName As String = Replace(oldName, matchingPattern, newPattern, , , CompareMethod.Text)
        Dim newFile As String = System.IO.Path.Combine(folderPath, newName & System.IO.Path.GetExtension(File))
        System.IO.File.Move(File, newFile)

        ' Open the new document
        Dim doc As Document = ThisApplication.Documents.Open(newFile, False)

        ' Rename the component occurrence
        If TypeOf doc Is PartDocument Then
            Dim partDoc As PartDocument = CType(doc, PartDocument)
            Dim partCompDef As PartComponentDefinition = partDoc.ComponentDefinition

            For Each occ As ComponentOccurrence In partCompDef.Occurrences
                Dim occName As String = occ.Name
                Dim newOccName As String = Replace(occName, matchingPattern, newPattern, , , CompareMethod.Text)
                occ.Rename(newOccName)
            Next

            partDoc.Save()
        ElseIf TypeOf doc Is AssemblyDocument Then
            Dim assyDoc As AssemblyDocument = CType(doc, AssemblyDocument)
            Dim assyCompDef As AssemblyComponentDefinition = assyDoc.ComponentDefinition

            For Each occ As ComponentOccurrence In assyCompDef.Occurrences
                Dim occName As String = occ.Name
                Dim newOccName As String = Replace(occName, matchingPattern, newPattern, , , CompareMethod.Text)
                occ.Rename(newOccName)
            Next

            assyDoc.Save()
        End If

        doc.Close()
    Next

    ' Refresh the file list
    ThisApplication.FileManager.RefreshAllDocuments()
End Sub
