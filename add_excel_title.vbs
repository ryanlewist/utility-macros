 'VB Script to add a title to excel tab
Option Explicit
 
Dim eapp, title_sheet, title_name, wkbk, Input_Excel, objArgs, lastColumn, ColLetter, firstRow, firstRowAddress

Input_Excel	 = WScript.Arguments(0) 'Full path and Input file name
title_sheet  = WScript.Arguments(1) 
title_name	 = WScript.Arguments(2) 

	
	Set objArgs = Wscript.Arguments
	Set eapp = CreateObject("Excel.Application")
	Set wkbk = eapp.Workbooks.Open(Input_Excel)
	eapp.Visible = false
	eapp.DisplayAlerts = false

	Const xlToRight = -4161 '-4161 is the value of xlToRight
	Const xlShiftDown = -4121 '-4121 is the value of xlShiftDown
	Const xlHAlignCenter = -4108 '-4108 is the value of xlHAlignCenter
	Const xlContinuous = 1 '1 is the value of xlContinous
	Const xlEdgeBottom = 9 '9 is the value of xlEdgeBottom
	Const xlEdgeLeft = 7 '7 is the value of xlEdgeLeft
	Const xlEdgeRight = 10 '10 is the value of xlEdgeRight
	Const xlEdgeTop = 8  '8 is the value of xlEdgeTop
	
	set lastColumn = wkbk.Worksheets(title_sheet).Range("A1").End(xlToRight)
	ColLetter = split(lastColumn.Address(1,0), "$")(0)

	wkbk.Worksheets(title_sheet).Activate
	set firstRow = wkbk.Worksheets(title_sheet).Range("A1").EntireRow
	firstRowAddress = split(firstRow.Address(1,0), "$")(0)
	firstRow.Insert(xlShiftDown)
	
	wkbk.Worksheets(title_sheet).Range("A2").copy wkbk.Worksheets(title_sheet).Range("A1")

	wkbk.Worksheets(title_sheet).Range("A1").Value = title_name
	with wkbk.Worksheets(title_sheet).Range("A1:" & ColLetter & 1)
		.HorizontalAlignment = xlHAlignCenter
		.WrapText = False
		.Orientation = 0
		.AddIndent = False
		.IndentLevel = 0
		.ShrinkToFit = False
		.Borders(xlEdgeLeft).LineStyle = xlContinuous
		.Borders(xlEdgeRight).LineStyle = xlContinuous
		.Borders(xlEdgeBottom).LineStyle = xlContinuous
		.Borders(xlEdgeTop).LineStyle = xlContinuous
		.Borders.Color = RGB(217,217,217)
		.merge
	end with
	
	wkbk.Worksheets(title_sheet).Cells.EntireColumn.Autofit
	
	wkbk.Worksheets(1).Activate
	wkbk.Save

eapp.Workbooks.Open(Input_Excel).Close
eapp.Quit
set wkbk = nothing