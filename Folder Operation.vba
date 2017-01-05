Sub folderOperation()

Dim fso As Object
Dim folderName As String
folderName = "C:\Users\" & Environ$("Username") & "\Desktop\Visteon Invoices\"
 
Set fso = CreateObject("Scripting.FileSystemObject")
 
If Not fso.FolderExists(folderName) Then

    fso.createFolder (folderName)
Else
    Set oFolder = fso.GetFolder(folderName)
    For Each ofile In oFolder.Files
        ofile.Delete True
    Next

End If

End Sub
