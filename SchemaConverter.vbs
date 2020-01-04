Sub StripLLP()
    Dim SrcFile As String
    Dim DestFile As String
    Dim Year As String: Year = "\2019-"
    SrcFile = ActiveWorkbook.Path & Year & "02-February.xlsx"
    DestFile = ActiveWorkbook.Path & Year & "February.xlsx"
    Dim Delimiter As String: Delimiter = "/"
    Dim SrcNoCol As Integer: SrcNoCol = 2
    Dim SrcNameCol As Integer: SrcNameCol = 3
    Dim SrcDateCol As Integer: SrcDateCol = 4
    Dim SrcStateCol As Integer: SrcStateCol = 13
    Dim SrcROCCol As Integer: SrcROCCol = 6
    Dim SrcPartnersCol As Integer: SrcPartnersCol = 12
    Dim SrcDPartnersCol As Integer: SrcDPartnersCol = 11
    Dim SrcObligationCol As Integer: SrcObligationCol = 10
    Dim SrcDivisionCol As Integer: SrcDivisionCol = 8
    Dim SrcActivityCol As Integer: SrcActivityCol = 9
    Dim DestCountCol As Integer: DestCountCol = 1
    Dim DestNoCol As Integer: DestNoCol = 2
    Dim DestNameCol As Integer: DestNameCol = 3
    Dim DestDateCol As Integer: DestDateCol = 4
    Dim DestStateCol As Integer: DestStateCol = 5
    Dim DestROCCol As Integer: DestROCCol = 6
    Dim DestPartnersCol As Integer: DestPartnersCol = 7
    Dim DestDPartnersCol As Integer: DestDPartnersCol = 8
    Dim DestObligationCol As Integer: DestObligationCol = 9
    Dim DestDivisionCol As Integer: DestDivisionCol = 10
    Dim DestActivityCol As Integer: DestActivityCol = 11
    Dim SrcBook As Workbook: Set SrcBook = Workbooks.Open(SrcFile)
    Dim DestBook As Workbook: Set DestBook = Workbooks.Open(DestFile)
    Dim SrcSheet As Worksheet: Set SrcSheet = SrcBook.Worksheets("Domestic LLP")
    Dim DestSheet As Worksheet: Set DestSheet = DestBook.Worksheets("Sheet1")
    Dim YearPos As Integer: YearPos = 2
    Dim MonthPos As Integer: MonthPos = 1
    Dim DayPos As Integer: DayPos = 0
    Dim i As Integer
    Dim cellvalue As String
    Dim delimited() As String
    Dim SrcRange As Range: Set SrcRange = SrcSheet.Range("A2", SrcSheet.Range("M2").End(xlDown))
    Dim DestRange As Range: Set DestRange = DestSheet.Range("A2", DestSheet.Range("K2").End(xlDown))
    NumRows = SrcSheet.Range("E2", SrcSheet.Range("E2").End(xlDown)).Rows.Count
    DestSheet.Range("A1:K1").Value = Array("Count", "No", "Name", "Date", "State", "ROC", "Partners", "Designated Partners", "Obligation", "Division", "Activity")
    For i = 1 To NumRows
        DestRange.Cells(RowIndex:=i, ColumnIndex:=DestCountCol).Value = "1"
        cellvalue = SrcRange.Cells(RowIndex:=i, ColumnIndex:=SrcDateCol).Value
        delimited = Split(Replace(cellvalue, ".", "/"), Delimiter)
        delimited = Split(Replace(cellvalue, "-", "/"), Delimiter)
        DestRange.Cells(RowIndex:=i, ColumnIndex:=DestDateCol).Value = delimited(YearPos) & "-" & delimited(MonthPos) & "-" & delimited(DayPos)
        DestRange.Cells(RowIndex:=i, ColumnIndex:=DestNoCol).Value = SrcRange.Cells(RowIndex:=i, ColumnIndex:=SrcNoCol).Value
        DestRange.Cells(RowIndex:=i, ColumnIndex:=DestNameCol).Value = SrcRange.Cells(RowIndex:=i, ColumnIndex:=SrcNameCol).Value
        DestRange.Cells(RowIndex:=i, ColumnIndex:=DestStateCol).Value = SrcRange.Cells(RowIndex:=i, ColumnIndex:=SrcStateCol).Value
        DestRange.Cells(RowIndex:=i, ColumnIndex:=DestROCCol).Value = SrcRange.Cells(RowIndex:=i, ColumnIndex:=SrcROCCol).Value
        DestRange.Cells(RowIndex:=i, ColumnIndex:=DestPartnersCol).Value = SrcRange.Cells(RowIndex:=i, ColumnIndex:=SrcPartnersCol).Value
        DestRange.Cells(RowIndex:=i, ColumnIndex:=DestDPartnersCol).Value = SrcRange.Cells(RowIndex:=i, ColumnIndex:=SrcDPartnersCol).Value
        DestRange.Cells(RowIndex:=i, ColumnIndex:=DestObligationCol).Value = SrcRange.Cells(RowIndex:=i, ColumnIndex:=SrcObligationCol).Value
        DestRange.Cells(RowIndex:=i, ColumnIndex:=DestDivisionCol).Value = SrcRange.Cells(RowIndex:=i, ColumnIndex:=SrcDivisionCol).Value
        'DestRange.Cells(RowIndex:=i, ColumnIndex:=DestDivisionCol).Value = "0"
        DestRange.Cells(RowIndex:=i, ColumnIndex:=DestActivityCol).Value = SrcRange.Cells(RowIndex:=i, ColumnIndex:=SrcActivityCol).Value
    Next
    DestSheet.Columns("F").Replace What:="ROC - ", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=False
    DestSheet.Columns("F").Replace What:="ROC-", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=False
    DestSheet.Columns("J").Replace What:="ZMCA_INDUS_ACT_", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    DestSheet.Columns("K").Replace What:="Industrial Activity - ", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
    DestBook.Close SaveChanges:=True
    SrcBook.Close SaveChanges:=False
End Sub
