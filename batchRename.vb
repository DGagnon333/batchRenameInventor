Sub Main
    ' Prompt the user for the matching pattern and new pattern
    Dim matchPattern As String = InputBox("Enter the matching pattern:")
    Dim newPattern As String = InputBox("Enter the new pattern:")

    ' Rename the files that match the pattern
    Dim currentDirectory As String = ThisDoc.Path
    Dim files As String() = System.IO.Directory.GetFiles(currentDirectory, "*.ipt").Concat(System.IO.Directory.GetFiles(currentDirectory, "*.iam")).ToArray()

    For Each file As String In files
        If System.IO.Path.GetFileNameWithoutExtension(file).Contains(matchPattern) Then
            Dim newFileName As String = System.IO.Path.Combine(currentDirectory, System.IO.Path.GetFileNameWithoutExtension(file).Replace(matchPattern, newPattern) & System.IO.Path.GetExtension(file))
            System.IO.File.Move(file, newFileName)
        End If
    Next

    ' Update all references to the renamed files in the active document
    Dim oDoc As Document = ThisDoc.Document
    oDoc.Update()

    ' Show a message box indicating the rename is complete
    MsgBox("Rename complete.")
End Sub
