vcsvFileName = WScript.Arguments.Item(0)
vxlsxFileName = WScript.Arguments.Item(1)
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(vcsvFileName)
objExcel.DisplayAlerts = False
objExcel.ActiveWorkbook.SaveAs vxlsxFileName, 51
objExcel.DisplayAlerts = True
objExcel.Workbooks(1).Save
objExcel.Workbooks(1).Close SaveChanges=True

objExcel.Quit

WScript.Quit
