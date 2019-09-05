Const xlToRight = -4161
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWB = objExcel.Workbooks.Open("C:\Users\chandu.s\Desktop\CARDAUTO\in\CARD.xlsx")
Set objSheet = objwb.Sheets("WORKING")
objSheet.Columns("N:N").Insert xlToRight
objSheet.Range("M:M").Select
objExcel.Selection.NumberFormat = "0.0000"
objSheet.Cells(7, 30).Value = objSheet.Cells(7, 12).Value + objSheet.Cells(7, 13).Value
objSheet.Cells(7, 14).Value = objSheet.Cells(7, 30).Value / 2
objSheet.Cells(8, 30).Value = objSheet.Cells(8, 12).Value + objSheet.Cells(8, 13).Value
objSheet.Cells(8, 14).Value = objSheet.Cells(8, 30).Value / 2
objSheet.Cells(9, 30).Value = objSheet.Cells(9, 12).Value + objSheet.Cells(9, 13).Value
objSheet.Cells(9, 14).Value = objSheet.Cells(9, 30).Value / 2
objSheet.Cells(10, 30).Value = objSheet.Cells(10, 12).Value + objSheet.Cells(10, 13).Value
objSheet.Cells(10, 14).Value = objSheet.Cells(10, 30).Value / 2
objSheet.Cells(11, 30).Value = objSheet.Cells(11, 12).Value + objSheet.Cells(11, 13).Value
objSheet.Cells(11, 14).Value = objSheet.Cells(11, 30).Value / 2
objSheet.Cells(12, 30).Value = objSheet.Cells(12, 12).Value + objSheet.Cells(12, 13).Value
objSheet.Cells(12, 14).Value = objSheet.Cells(12, 30).Value / 2
objSheet.Cells(13, 30).Value = objSheet.Cells(13, 12).Value + objSheet.Cells(13, 13).Value
objSheet.Cells(13, 14).Value = objSheet.Cells(13, 30).Value / 2
objSheet.Cells(14, 30).Value = objSheet.Cells(14, 12).Value + objSheet.Cells(14, 13).Value
objSheet.Cells(14, 14).Value = objSheet.Cells(14, 30).Value / 2
objSheet.Cells(15, 30).Value = objSheet.Cells(15, 12).Value + objSheet.Cells(15, 13).Value
objSheet.Cells(15, 14).Value = objSheet.Cells(15, 30).Value / 2
objSheet.Cells(16, 30).Value = objSheet.Cells(16, 12).Value + objSheet.Cells(16, 13).Value
objSheet.Cells(16, 14).Value = objSheet.Cells(16, 30).Value / 2
objSheet.Cells(17, 30).Value = objSheet.Cells(17, 12).Value + objSheet.Cells(17, 13).Value
objSheet.Cells(17, 14).Value = objSheet.Cells(17, 30).Value / 2
objSheet.Cells(18, 30).Value = objSheet.Cells(18, 12).Value + objSheet.Cells(18, 13).Value
objSheet.Cells(18, 14).Value = objSheet.Cells(18, 30).Value / 2
objSheet.Cells(19, 30).Value = objSheet.Cells(19, 12).Value + objSheet.Cells(19, 13).Value
objSheet.Cells(19, 14).Value = objSheet.Cells(19, 30).Value / 2

objSheet.Cells(1, 8).Value = "   AUTOMATED BY CHANDU SANJITH T , HAPPY COADING!!!   "
objSheet.Cells(1, 8).Interior.ColorIndex = 48
objSheet.Cells(1, 9).Interior.ColorIndex = 48
objSheet.Cells(1, 10).Interior.ColorIndex = 48
objSheet.Cells(1, 11).Interior.ColorIndex = 48
objSheet.Cells(1, 12).Interior.ColorIndex = 48
objSheet.Cells(1, 13).Interior.ColorIndex = 48
objSheet.Cells(1, 8).Font.ColorIndex = 9

objSheet.Cells(7, 14).Select
objExcel.Selection.NumberFormat = "0.0000"
objSheet.Cells(8, 14).Select
objExcel.Selection.NumberFormat = "0.0000"
objSheet.Cells(9, 14).Select
objExcel.Selection.NumberFormat = "0.0000"
objSheet.Cells(10, 14).Select
objExcel.Selection.NumberFormat = "0.0000"
objSheet.Cells(11, 14).Select
objExcel.Selection.NumberFormat = "0.0000"
objSheet.Cells(12, 14).Select
objExcel.Selection.NumberFormat = "0.0000"
objSheet.Cells(13, 14).Select
objExcel.Selection.NumberFormat = "0.0000"
objSheet.Cells(14, 14).Select
objExcel.Selection.NumberFormat = "0.0000"
objSheet.Cells(15, 14).Select
objExcel.Selection.NumberFormat = "0.0000"
objSheet.Cells(16, 14).Select
objExcel.Selection.NumberFormat = "0.0000"
objSheet.Cells(17, 14).Select
objExcel.Selection.NumberFormat = "0.0000"
objSheet.Cells(18, 14).Select
objExcel.Selection.NumberFormat = "0.0000"
objSheet.Cells(19, 14).Select
objExcel.Selection.NumberFormat = "0.0000"

objSheet.Cells(13, 30).Value = objSheet.Cells(13, 5).Value + objSheet.Cells(13, 6).Value + objSheet.Cells(13, 7).Value + objSheet.Cells(13, 8).Value + objSheet.Cells(13, 9).Value + objSheet.Cells(13, 10).Value 
objSheet.Cells(13, 11).Value = objSheet.Cells(13, 30).Value / 6
objSheet.Cells(15, 30).Value = objSheet.Cells(15, 5).Value + objSheet.Cells(15, 6).Value + objSheet.Cells(15, 7).Value + objSheet.Cells(15, 8).Value + objSheet.Cells(15, 9).Value + objSheet.Cells(15, 10).Value 
objSheet.Cells(15, 11).Value = objSheet.Cells(15, 30).Value / 6
objSheet.Cells(17, 30).Value = objSheet.Cells(17, 5).Value + objSheet.Cells(17, 6).Value + objSheet.Cells(17, 7).Value + objSheet.Cells(17, 8).Value + objSheet.Cells(17, 9).Value + objSheet.Cells(17, 10).Value 
objSheet.Cells(17, 11).Value = objSheet.Cells(17, 30).Value / 6

objSheet.Cells(13, 11).Select
objExcel.Selection.NumberFormat = "0.0000"
objSheet.Cells(15, 11).Select
objExcel.Selection.NumberFormat = "0.0000"
objSheet.Cells(17, 11).Select
objExcel.Selection.NumberFormat = "0.0000"

objSheet.Cells(13, 11).Interior.ColorIndex = 44
objSheet.Cells(13, 10).Interior.ColorIndex = 44
objSheet.Cells(13, 9).Interior.ColorIndex = 44
objSheet.Cells(13, 8).Interior.ColorIndex = 44
objSheet.Cells(13, 7).Interior.ColorIndex = 44
objSheet.Cells(13, 6).Interior.ColorIndex = 44
objSheet.Cells(13, 5).Interior.ColorIndex = 44
objSheet.Cells(13, 4).Interior.ColorIndex = 44
objSheet.Cells(13, 3).Interior.ColorIndex = 44
objSheet.Cells(13, 2).Interior.ColorIndex = 44

objSheet.Cells(15, 11).Interior.ColorIndex = 44
objSheet.Cells(15, 10).Interior.ColorIndex = 44
objSheet.Cells(15, 9).Interior.ColorIndex = 44
objSheet.Cells(15, 8).Interior.ColorIndex = 44
objSheet.Cells(15, 7).Interior.ColorIndex = 44
objSheet.Cells(15, 6).Interior.ColorIndex = 44
objSheet.Cells(15, 5).Interior.ColorIndex = 44
objSheet.Cells(15, 4).Interior.ColorIndex = 44
objSheet.Cells(15, 3).Interior.ColorIndex = 44
objSheet.Cells(15, 2).Interior.ColorIndex = 44

objSheet.Cells(17, 11).Interior.ColorIndex = 44
objSheet.Cells(17, 10).Interior.ColorIndex = 44
objSheet.Cells(17, 9).Interior.ColorIndex = 44
objSheet.Cells(17, 8).Interior.ColorIndex = 44
objSheet.Cells(17, 7).Interior.ColorIndex = 44
objSheet.Cells(17, 6).Interior.ColorIndex = 44
objSheet.Cells(17, 5).Interior.ColorIndex = 44
objSheet.Cells(17, 4).Interior.ColorIndex = 44
objSheet.Cells(17, 3).Interior.ColorIndex = 44
objSheet.Cells(17, 2).Interior.ColorIndex = 44
objWB.Close True
objExcel.Quit	