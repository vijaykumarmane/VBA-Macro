' One one pdf creater
Sub pdfCreater()

Dim counter, posInString As Integer

For counter = 1 To Worksheets.count
    nameOfsheet = Worksheets(counter).Name
    Status = Sheets(Name).Visible
    If Status = -1 Then
    Worksheets(counter).ExportAsFixedFormat _
    Type:=xlTypePDF, _
    Filename:="C:\Users\" & Environ$("Username") & "\Desktop\Visteon Invoices\" & Worksheets(counter).Name, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=False
    End If
Next

End Sub


' Consolidation of sheets to one pdf
Sub singlePdfWorkbook()

Dim counter As Integer
sheetCount = 0

Sheets("Billing Template").Activate

For Each Sheet In Worksheets
        sheetName = Sheet.Name
        Status = Sheets(sheetName).Visible
        If Status = -1 Then
            Sheets(sheetName).Select Replace:=False
        End If
Next Sheet

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
    "C:\Users\vijaykumar.mane\Desktop\PDF Sheets\tempo.pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, _
     IgnorePrintAreas:=False, OpenAfterPublish:=False

End Sub
