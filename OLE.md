# Remove All OLE Links
https://knowledge.autodesk.com/support/inventor/troubleshooting/caas/sfdcarticles/sfdcarticles/Inventor-have-OLE-file-can-not-be-deleted.html
```vba
Public Sub DeleteOLEReference()
    Dim oDoc As Document
    Set oDoc = ThisApplication.ActiveDocument

    If oDoc.ReferencedOLEFileDescriptors.Count = 0 Then
        MsgBox "There aren't any OLE references in this document."
        Exit Sub
    End If

    Dim aOLERefs() As ReferencedOLEFileDescriptor
    ReDim aOLERefs(oDoc.ReferencedOLEFileDescriptors.Count - 1)

    Dim iRefCount As Integer
    iRefCount = oDoc.ReferencedOLEFileDescriptors.Count
    Dim i As Integer
    For i = 1 To iRefCount
        Set aOLERefs(i - 1) = oDoc.ReferencedOLEFileDescriptors.Item(i)
    Next

    For i = 1 To iRefCount
        If MsgBox("Delete """ & aOLERefs(i - 1).FullFileName & """?", vbQuestion + vbYesNo) = vbYes Then
            aOLERefs(i - 1).Delete
        End If
    Next
End Sub
```

#
```vba
```
