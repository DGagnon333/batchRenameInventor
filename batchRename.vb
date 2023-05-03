Sub Main
    Dim dlg As New Inventor.SimpleDialog("Batch Rename")
    Dim txtMatch As Inventor.TextBox = dlg.TextBox("txtMatch", "Matching Pattern:")
    Dim txtNew As Inventor.TextBox = dlg.TextBox("txtNew", "New Pattern:")
    Dim btnRename As Inventor.Button = dlg.Button("btnRename", "Rename")

    dlg.Show()

    If dlg.DialogCancelled = False And btnRename.Activated = True Then
        Dim currentDirectory As String = ThisDoc.Path

        Dim files As String() = System.IO.Directory.GetFiles(currentDirectory, "*.ipt").Concat(System.IO.Directory.GetFiles(currentDirectory, "*.iam")).ToArray()

        For Each file As String In files
            If System.IO.Path.GetFileNameWithoutExtension(file).Contains(txtMatch.Value) Then
                Dim newFileName As String = System.IO.Path.Combine(currentDirectory, System.IO.Path.GetFileNameWithoutExtension(file).Replace(txtMatch.Value, txtNew.Value) & System.IO.Path.GetExtension(file))
                System.IO.File.Move(file, newFileName)
            End If
        Next

        Dim oDoc As Document = ThisDoc.Document
        oDoc.Update()
    End If
End Sub
