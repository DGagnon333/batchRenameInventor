Sub main()
    ' Get the active document
    Dim activeDoc As Document = ThisApplication.ActiveDocument

    ' Check if the document is a part or assembly file
    If Not TypeOf activeDoc Is PartDocument And Not TypeOf activeDoc Is AssemblyDocument Then
        MsgBox("This script is intended to be used with part and assembly documents only.")
        Exit Sub
    End If

    ' Get the matching and new patterns
    Dim matchingPattern As String = InputBox("Enter the matching pattern:")
    If matchingPattern = "" Then
        MsgBox("Matching pattern cannot be empty.")
        Exit Sub
    End If

    Dim newPattern As String = InputBox("Enter the new pattern:")
    If newPattern = "" Then
        MsgBox("New pattern cannot be empty.")
        Exit Sub
    End If

    ' Rename the components in the active document
    If TypeOf activeDoc Is PartDocument Then
        Dim partDoc As PartDocument = CType(activeDoc, PartDocument)
        RenamePartComponents(partDoc, matchingPattern, newPattern)
    Else
        Dim asmDoc As AssemblyDocument = CType(activeDoc, AssemblyDocument)
        RenameAssemblyComponents(asmDoc, matchingPattern, newPattern)
    End If

    ' Display a message when renaming is complete
    MsgBox("Renaming complete.")
End Sub

Sub RenameAssemblyComponents(doc As AssemblyDocument, matchingPattern As String, newPattern As String)
    Dim compOccs As IEnumerable(Of ComponentOccurrence) = doc.ComponentDefinition.Occurrences.OfType(Of ComponentOccurrence)()

    For Each compOcc As ComponentOccurrence In compOccs
        Dim occName As String = compOcc.Name
        Dim newOccName As String = Replace(occName, matchingPattern, newPattern, , , CompareMethod.Text)
        compOcc.Name = newOccName
    Next

    ' Recursively rename occurrences in referenced documents
    For Each refDoc As Document In doc.AllReferencedDocuments
        If TypeOf refDoc Is File Then
            Dim file As File = refDoc.File
            If file.FileType = FileTypeEnum.kAssemblyFileType Then
                RenameAssemblyComponents(refDoc, matchingPattern, newPattern)
            ElseIf file.FileType = FileTypeEnum.kPartFileType Then
                RenamePartComponents(refDoc, matchingPattern, newPattern)
            End If
        End If
    Next
End Sub

Sub RenamePartComponents(doc As PartDocument, matchingPattern As String, newPattern As String)
    Dim compOccs As IEnumerable(Of ComponentOccurrence) = doc.ComponentDefinition.Occurrences.OfType(Of ComponentOccurrence)()

    For Each compOcc As ComponentOccurrence In compOccs
        Dim occName As String = compOcc.Name
        Dim newOccName As String = Replace(occName, matchingPattern, newPattern, , , CompareMethod.Text)
        compOcc.Name = newOccName
    Next

    ' Recursively rename occurrences in referenced documents
    For Each refDoc As Document In doc.AllReferencedDocuments
        If TypeOf refDoc Is File Then
            Dim file As File = refDoc.File
            If file.FileType = FileTypeEnum.kAssemblyFileType Then
                RenameAssemblyComponents(refDoc, matchingPattern, newPattern)
            ElseIf file.FileType = FileTypeEnum.kPartFileType Then
                RenamePartComponents(refDoc, matchingPattern, newPattern)
            End If
        End If
    Next
End Sub
