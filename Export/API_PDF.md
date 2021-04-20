# API export to PDF

```vba
Public Sub PublishPDF()
    ' Get the PDF translator Add-In.
    Dim PDFAddIn As TranslatorAddIn
    Set PDFAddIn = ThisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")

    'Set a reference to the active document (the document to be published).
    Dim oDocument As Document
    Set oDocument = ThisApplication.ActiveDocument

    Dim oContext As TranslationContext
    Set oContext = ThisApplication.TransientObjects.CreateTranslationContext
    oContext.Type = kFileBrowseIOMechanism

    ' Create a NameValueMap object
    Dim oOptions As NameValueMap
    Set oOptions = ThisApplication.TransientObjects.CreateNameValueMap

    ' Create a DataMedium object
    Dim oDataMedium As DataMedium
    Set oDataMedium = ThisApplication.TransientObjects.CreateDataMedium

    ' Check whether the translator has 'SaveCopyAs' options
    If PDFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then

        ' Options for drawings...

        oOptions.Value("All_Color_AS_Black") = 0

        'oOptions.Value("Remove_Line_Weights") = 0
        'oOptions.Value("Vector_Resolution") = 400
        'oOptions.Value("Sheet_Range") = kPrintAllSheets
        'oOptions.Value("Custom_Begin_Sheet") = 2
        'oOptions.Value("Custom_End_Sheet") = 4

    End If

    'Set the destination file name
    oDataMedium.FileName = "c:\temp\test.pdf"

    'Publish document.
    Call PDFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
End Sub
```
