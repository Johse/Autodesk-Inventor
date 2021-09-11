´´´vba
''This macro wil delete unresolved constraints.

''Macro created by Stefaan Boel
''Copyright by Inventor Wizard (http://www.inventorwizard.be)
''Use this macro at your own risk.
''You may only copy/modify this or part of the code if you leave this header!

Public Sub DeleteUnresolvedConstraints()

Dim oAssDoc As AssemblyDocument
Set oAssDoc = ThisApplication.ActiveDocument

Dim oConstraint As AssemblyConstraint

Dim iQuestion1 As Integer
iQuestion1 = MsgBox("Delete all constraint?", vbYesNo) 'Do you whant to delete all the constraints or not?


For Each oConstraint In oAssDoc.ComponentDefinition.Constraints
If oConstraint.HealthStatus = kInconsistentHealth Then

'Delete all constraints without asking
If iQuestion1 = vbYes Then
    oConstraint.Delete

Else

'Delete constraints one by one, only when selected Yes.
Dim iQuestion2 As Integer
iQuestion2 = MsgBox("Unresolved constraint found: '" & oConstraint.Name & "'. Delete constraint?", vbYesNo)
If iQuestion2 = vbYes Then
oConstraint.Delete

End If
End If
End If

Next oConstraint
MsgBox "Done deleting constraints!", vbInformation
End Sub
