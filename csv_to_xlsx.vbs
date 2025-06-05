folderPath = InputBox("Enter the path to the folder (without filename):", "CSV to XLSX")
fileName = InputBox("Enter the filename (without .csv):", "CSV to XLSX")

csvFile = folderPath & "\" & fileName & ".csv"
xlsxFile = folderPath & "\" & fileName & ".xlsx"

Set xl = CreateObject("Excel.Application")
xl.Visible = False

On Error Resume Next
Set wb = xl.Workbooks.Open(csvFile)
If Err.Number <> 0 Then
    MsgBox "Error opening file: " & csvFile, vbCritical, "Error"
    xl.Quit
    WScript.Quit
End If
On Error GoTo 0

wb.SaveAs xlsxFile, 51  ' 51 = xlOpenXMLWorkbook (.xlsx)
wb.Close False
xl.Quit

MsgBox "Successfully converted: " & xlsxFile, vbInformation, "Done"
