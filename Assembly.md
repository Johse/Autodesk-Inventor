# Copy all assembly user parameters to contained Parts
https://forums.autodesk.com/t5/inventor-customization/ilogic-push-all-assembly-user-parameters-to-parts/td-p/7961847

```vba
Public Sub Main()
	CopyUserParams()
End Sub

Private Sub CopyUserParams()
    If ThisDoc.Document.DocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
        MsgBox("The active document must be an assembly.")
        Return
    End If

    Dim asmDoc As Inventor.AssemblyDocument = ThisDoc.Document	
    For Each refDoc As Inventor.Document In asmDoc.AllReferencedDocuments
        ' Look for part documents.
        If refDoc.DocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject Then
            Dim partDoc As Inventor.PartDocument = refDoc
            Dim refDocUserParams As UserParameters = partDoc.ComponentDefinition.Parameters.UserParameters

            ' Add the assembly parameters to the part.
            For Each asmUserParam As UserParameter In asmDoc.ComponentDefinition.Parameters.UserParameters
                ' Check to see if the parameter already exists.
                Dim checkParam As UserParameter = Nothing
                Try
                    checkParam = refDocUserParams.Item(asmUserParam.Name)
                Catch ex As Exception
                    checkParam = Nothing
                End Try
				
				Dim svalue As String
            	Dim bvalue As Boolean
				
                If checkParam Is Nothing Then
                   ' Create the missing parameter.
                If asmUserParam.Units = "Text" Then
                    svalue = Replace(asmUserParam.Value, Chr(34), "")
                    Call refDocUserParams.AddByValue(asmUserParam.Name, svalue, asmUserParam.Units)
                ElseIf asmUserParam.Units = "Boolean" Then
                    bvalue = Replace(asmUserParam.Value, Chr(34), "")
                    Call refDocUserParams.AddByValue(asmUserParam.Name, CBool(bvalue), asmUserParam.Units)
                Else
                    Call refDocUserParams.AddByExpression(asmUserParam.Name, asmUserParam.Expression, asmUserParam.Units)
                End If
            Else
                ' Update the value of the existing parameter.
                If asmUserParam.Units = "Text" Then
                    svalue = Replace(asmUserParam.Value, Chr(34), "")
                    checkParam.Value = svalue
                ElseIf asmUserParam.Units = "Boolean" Then
                    bvalue = Replace(asmUserParam.Value, Chr(34), "")
                    checkParam.Value = bvalue
                Else
                    checkParam.Expression = asmUserParam.Expression
                End If
            End If
            Next
        End If
    Next
End Sub
```
