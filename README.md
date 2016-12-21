# EXCEL-VBA-sort-for-every-sheets-in-a-workbook
#It's just suit for one  situation of sort in EXCEL.It's not a general version.
Sub 排序排序()
'
' 排序排序 宏
'
'
  For i = 1 To Sheets.Count
        Sheets(i).Select
         Range("A3:G9").Select
        Range(Selection, Selection.End(xlDown)).Select
        ActiveWorkbook.Sheets(i).Sort.SortFields.Clear
        ActiveWorkbook.Sheets(i).Sort.SortFields.Add Key:=Range("C3:C562"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Sheets(i).Sort.SortFields.Add Key:=Range("A3:A562"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        "15t16,17t18,19,20,21t22,23,24,25,26,27t28,29,30t33,34t35,36t37", DataOption _
        :=xlSortNormal
      With ActiveWorkbook.Sheets(i).Sort
        .SetRange Range("A3:G562")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
      End With
    Next
End Sub
