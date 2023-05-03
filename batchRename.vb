Sub Main
    Dim matchingPattern As String = InputBox("Enter the matching pattern:")
    Dim newPattern As String = InputBox("Enter the new pattern:")

    Dim doc As Document = ThisDoc.Document

    For Each file As Document In doc.AllReferencedDocuments
        Dim fileName As String = file.FullFileName
        Dim fileDir As String = System.IO.Path.GetDirectoryName(fileName)
        Dim fileExt As String = System.IO.Path.GetExtension(fileName)
        Dim newName As String = System.IO.Path.GetFileNameWithoutExtension(fileName)

        If newName.Contains(matchingPattern) Then
            newName = newName.Replace(matchingPattern, newPattern)
            newName = System.IO.Path.Combine(fileDir, newName + fileExt)
            System.IO.File.Move(fileName, newName)
            file.FullFileName = newName
        End If
    Next
End Sub
