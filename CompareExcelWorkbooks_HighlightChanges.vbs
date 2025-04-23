Option Explicit

Dim objExcel, wbOld, wbNew
Dim obj, scriptPath
Dim wsOld, wsNew
Dim sheetNameDict, sheetName
Dim lastRow, lastCol
Dim i, j

Set obj = CreateObject("Scripting.FileSystemObject")
scriptPath = obj.GetParentFolderName(WScript.ScriptFullName)

' Set the paths to your old and new Excel files
oldFile = scriptPath & "\old.xlsx"
newFile = scriptPath & "\new.xlsx"

' Create Excel application
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

' Open both workbooks
Set wbOld = objExcel.Workbooks.Open(oldFile)
Set wbNew = objExcel.Workbooks.Open(newFile)

' Create dictionary with the old report's sheets
Set sheetNameDict = CreateObject("Scripting.Dictionary")

' Populate dictionary with the old report's sheets
For i = 1 To wbOld.Sheets.Count
	sheetNameDict.Add wbOld.Sheets(i).Name, wbOld.Sheets(i)
Next

' Loop through the new report's sheets and compare only matching ones
For i = 1 To wbNew.Sheets.Count
	sheetName = wbNew.Sheets(i).Name
	sheetNameDict.Exists(sheetName) Then
		Set wsNew = wbNew.Sheets(sheetName)
		Set wsOld = sheetNameDict(sheetName)
		
		' Determine last used row and column (new report's sheet)
		lastRow = wsNew.UsedRange.Rows.Count
		lastCol = wsNew.UsedRange.Columns.Count

		' Compare cells
		For j = 1 To lastRow
			Dim k
			For k = 1 To lastCol
				If wsNew.Cells(j, k).Value <> wsOld.Cells(j, k).Value Then
					wsNew.Cells(j, k).Interior.Color = RGB(255, 255, 0)
				End If
			Next
		Next
	End If
Next

' Save and close the updated new workbook
wbNew.Save
wbNew.Close

' Close the old workbook
wbOld.Close False

' Quit Excel
objExcel.Quit
objExcel.Application.Quit

' Clean up
Set wsOld = Nothing
Set wsNew = Nothing
Set wbOld = Nothing
Set wbNew = Nothing
Set objExcel = Nothing

MsgBox "Comparison complete. Changes are highlighted in all sheets in the new Report.", vbInformation