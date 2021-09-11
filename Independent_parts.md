''This macro get all parts out of a pattern and ground them
''The pattern itself will be deleted.

''Macro created by Teun Ham & Dion Snoeijen
''Copyright by Inventor Wizard (http://www.inventorwizard.be)
''Use this macro at your own risk.
''You may only copy/modify this or part of the code if you leave this header!

Public Sub Make_Component_Pattern_Independent()

 Dim oDoc As Document
    
    Set oDoc = ThisApplication.ActiveDocument
    If oDoc.DocumentType <> kAssemblyDocumentObject Then
        MsgBox "Het actieve document moet een assembly document zijn (.iam)"
        Exit Sub
    End If
        MakeComponentPatternIndependent
End Sub

Sub MakeComponentPatternIndependent()

Dim oAssDoc As AssemblyDocument
Dim oPattern As OccurrencePattern
Dim oPatternElement As OccurrencePatternElement
Dim oOccurence As ComponentOccurrence

Set oAssDoc = ThisApplication.ActiveDocument
'Traverse all the patterns in the assembly:
For Each oPattern In oAssDoc.ComponentDefinition.OccurrencePatterns
    'Traverse all elements in the pattern to make them independent:
    For Each oPatternElement In oPattern.OccurrencePatternElements
        'Traverse all occurences in the element in order to Groud them:
        For Each oOccurence In oPatternElement.Occurrences
            oOccurence.Grounded = True
        Next oOccurence
        On Error Resume Next
        oPatternElement.Independent = True
        If Err Then
            Err.Clear
            End If
    Next oPatternElement
Next oPattern

'Delete all patterns in the assembly
For Each oPattern In oAssDoc.ComponentDefinition.OccurrencePatterns
    oPattern.Delete
Next oPattern
End Sub
