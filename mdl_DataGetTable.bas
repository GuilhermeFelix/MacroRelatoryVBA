Attribute VB_Name = "mdl_DataGetTable"
Function TemplateRelatoryMount(ByRef wsData As Worksheet, ByRef wsOut As Worksheet)

    On Error GoTo ErrChk

    errRet = 1
    
    Dim wb As Workbook
    Dim tbitems As ListObject
    
    Set wb = ThisWorkbook
        
    On Error Resume Next
    Sheets("MASTER").Select
    
     
    strTbCols = "Item;Cód;Descrição;Unid.;Prev. Entr.;Qt. Prev.;Conv.;Vl. Unit.;" & _
                "% D;% IPI;D. Total;Vl. Total"
    aTbCols = Split(strTbCols, ";")
    
    strTbColsFormat = ";@;@;@;dd/mm/yyyy;0.00;0.00;$ #,##0.00;0.00%;0.00%;$ #,##0.00;$ #,##0.00"
    aTbColsFormat = Split(strTbColsFormat, ";")
    
    wsOut.Activate
    
    Set tbitems = wsOut.ListObjects.Add(xlSrcRange, wsOut.Range("$A$5:$AC$5"), , xlNo) '("$A$5:$" & Chr(64 + UBound(aTbCols) + 1) & "$5"), , xlNo)
    tbitems.Name = "tbItems"
    
    For i = 0 To UBound(aTbCols)
        tbitems.ListColumns(i + 1).Name = aTbCols(i)
        tbitems.ListColumns(i + 1).Range.NumberFormat = aTbColsFormat(i)
    Next
    
    wsData.Activate
    
    l = 2
    lin = 0
    While wsData.Range("A" & l).Value <> ""
                
                'Example:
                'Intersect(tbitems.ListRows(lin).Range, tbitems.ListColumns(aTbCols(1)).Range) =
                'Intersect(tbitems.ListRows(lin).Range, tbitems.ListColumns(aTbCols(2)).Range) =
                'Intersect(tbitems.ListRows(lin).Range, tbitems.ListColumns(aTbCols(3)).Range) =
                'Intersect(tbitems.ListRows(lin).Range, tbitems.ListColumns(aTbCols(4)).Range) =
            
           
        
        l = l + 1
    Wend
    
    wsOut.Activate
    wsOut.Range("A1").Activate
    
    
    
    'Save
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim NovoWB As Workbook
    Set NovoWB = Workbooks.Add(xlWBATWorksheet)
    With NovoWB
        ThisWorkbook.ActiveSheet.Copy After:=.Worksheets(.Worksheets.Count)
        .Worksheets(1).Delete
        .SaveAs ThisWorkbook.Path & "\Novo Arquivo.xlsx"
        .Close False
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
       
    Exit Function

ErrChk:
    MsgBox Err.Description
    
    Stop
    Resume

End Function


