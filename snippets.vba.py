Sub copy_chart_to_powerpoint_from(chart_record As ChartRecord)
    Dim chart_ As ChartObject
    Dim slide_ As PowerPoint.Slide
    Dim shape_ As PowerPoint.ShapeRange
    
    Set chart_ = get_chart(sheet_name:=chart_record.sheet, name:=chart_record.name)
    Set slide_ = pwrpnt_pres.Slides(Int(chart_record.slide_number))
    
    'The cutcopymode and activate mess is required to avoid
    'a Runtime error '1004' Application-defined or object-defined error
    'Pure black magic. Unknown why this works.
    Application.CutCopyMode = False
    chart_.Activate
    chart_.Copy
    Application.CutCopyMode = False
    Set shape_ = slide_.Shapes.PasteSpecial(DataType:=ppPastePNG)
    shape_.LockAspectRatio = False
    shape_.height = chart_record.height
    shape_.width = chart_record.width
    shape_.left = chart_record.left
    shape_.top = chart_record.top    
End Sub

Function getSheetWithDefault(name As String, Optional wb As Excel.Workbook) As Excel.Worksheet
        If wb Is Nothing Then
            Set wb = ThisWorkbook
        End If

        If Not sheetExists(name, wb) Then
            wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)).name = name
        End If

        Set getSheetWithDefault = wb.Sheets(name)
End Function

Function sheetExists(name As String, Optional wb As Excel.Workbook) As Boolean
    Dim sheet As Excel.Worksheet

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If

    sheetExists = False
    For Each sheet In wb.Worksheets
        If sheet.name = name Then
            sheetExists = True
            Exit Function
        End If
    Next sheet
End Function