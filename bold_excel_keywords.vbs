 'VB Script to add a title to excel tab
Option Explicit
 
Dim eapp, title_sheet, keyword, wkbk, Input_Excel, objArgs, firstFound,FSO, FoundCell

Input_Excel	 = WScript.Arguments(0) 'Full path and Input file name
title_sheet  = WScript.Arguments(1) 
keyword	 = WScript.Arguments(2)

	
	Set objArgs = Wscript.Arguments
	Set eapp = CreateObject("Excel.Application")
	Set wkbk = eapp.Workbooks.Open(Input_Excel)
	eapp.Visible = false
	eapp.DisplayAlerts = false

	wkbk.Worksheets(title_sheet).Activate
	Const xlWhole = 1

	Set FSO = CreateObject("Scripting.FileSystemObject")

	Set FoundCell = wkbk.Worksheets(title_sheet).usedRange.Find(keyword, , , xlWhole)
	If Not FoundCell Is Nothing Then
	  FoundCell.EntireRow.Font.Bold = True
	  firstFound = FoundCell.Address
	  Do 
		set FoundCell = wkbk.Worksheets(title_sheet).usedRange.FindNext(FoundCell)
		FoundCell.EntireRow.Font.Bold = True
		Loop While FoundCell.Address <> firstFound
	End If
	
	wkbk.Worksheets(title_sheet).Cells.EntireColumn.Autofit
	
	wkbk.Worksheets(1).Activate
	wkbk.Save

eapp.Workbooks.Open(Input_Excel).Close
eapp.Quit
set wkbk = nothing

