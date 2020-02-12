' Date: 110120, Time:01:56PM
' Author: Krishna M
' Objective: Creating an excel file for a given path


'#--------------------------------Creating Excel Object (Start)------------------------------#'

Set objExcel = CreateObject("Excel.Application")

'#--------------------------------Coding Conventions (End)-----------------------------------#'




'#--------------------------------Coding Convenrtions (Start)--------------------------------#'

' use indentation inside the function
' write following information before creating a function
' What does function do?
' what are the inputs?
' What are the outputs?

'#--------------------------------Coding Conventions (End)-----------------------------------#'



'#--------------------------------Function Declaration (Start)-------------------------------#'

'##---------------------Function 1--------------------------##
'Function to open excel file for a given path
'Input path of the Excel file
'output Visible Excel file


function OPEN_EXCEL(filepath)
	objExcel.Workbooks.Open(filepath)
	objExcel.Application.Visible = True
end function



' function OPEN_EXCEL(filepath)
	' MsgBox("In side function||")
	' Set objWoobjExcel.Workbooks.Open(filepath) 
	' MsgBox("File opened||||")
	' 'Set objWorksheet = objWorkbook.Worksheets(1)
	' objExcel.Application.Visible = True
	' MsgBox("Workbook returned ||")
' end function




'##-------------------End of Function 1--------------------##'




'##---------------------Function 2--------------------------##

function  FIND_DATA(filepath)

	Set obj1 = objExcel.Workbooks.Open(filepath)
	objExcel.Application.Visible = True

	Set objSheet = obj1.Worksheets("Sheet1")
	
	' Get the topmost row that has data
	nTop = objSheet.UsedRange.Row

	' Get leftmost column that has data
	nLeft = objSheet.UsedRange.Column
	
	' Get the number of used rows
	nUsedRows = objSheet.UsedRange.Rows.Count
	
	
	' Get the number of used columns
	nUsedCols = objSheet.UsedRange.Columns.Count
	
	' Get the lowest row that has data
	nBottom = nTop + nUsedRows - 1
	
	' Get the rightmost column that has data
	nRight = nLeft + nUsedCols - 1
	
	
	MsgBox("Top Row = " & nTop  & "  Left most Column = " & nLeft)
	MsgBox("Bottom Row = " & nBottom  & "  Right most Column = " & nRight)
	
	
	'Number of columns
	numCols = objSheet.UsedRange.Columns.count
	
	MsgBox("Number of Columns = " & numCols)
	'MsgBox "Entered value is " & Range("A1").Value
	
	' Dim MyArray() As Variant 
	' MyArray() = Array("Superhero", "City", "State", "Country", "Publisher", "Demographics", _
    ' "Planet", "Flying Abilities", "Vehicle", "Sidekick", "Powers")
	' For i = 0 To numCols
		' MyArray(i) = .Cells(1, 1 + i).Value
	' Next i
	
	' for i = 0 to uBound(myArray())
	' debug.print myArray(i)
	' Next i

	objExcel.Application.Visible = True
end function
'##-------------------End of Function 2--------------------##'


FIND_DATA("C:\\Users\\amit2\\Desktop\\Dreawe[12FEB2020]\\xlAutoLibrary-master\\xlAutoLibrary\\Input\\In01.xlsx")









