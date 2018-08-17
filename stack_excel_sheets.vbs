 'VB Script to copy tables from different worksheet to single worksheet. Deletes the copied sheet if requested
Option Explicit
 
Dim eapp, copy_wksht, paste_loc, wkbk, paste_wksht, copy_rnge, paste_rnge, Input_Excel, objArgs, veryLastCellCopy, LastCellCopyAddress, veryLastCellPaste, LastColLetterPaste, LastRowNumPaste,del_sheet, vert_offset, horz_offset

Input_Excel = WScript.Arguments(0) 'Full path and Input file name
copy_wksht  = WScript.Arguments(1) 
paste_wksht = WScript.Arguments(2) 
paste_loc	= WScript.Arguments(3) 
del_sheet	= WScript.Arguments(4)
vert_offset = WScript.Arguments(5)
horz_offset = WScript.Arguments(6)
	
	Set objArgs = Wscript.Arguments
	Set eapp = CreateObject("Excel.Application")
	Set wkbk = eapp.Workbooks.Open(Input_Excel)
	eapp.Visible = false
	eapp.DisplayAlerts = false

	Const xlCellTypeLastCell = 11 '11 is the value of xlCellTypeLastCell
	
	wkbk.Worksheets(copy_wksht).Activate
	set veryLastCellCopy = wkbk.Worksheets(copy_wksht).UsedRange

	LastCellCopyAddress = veryLastCellCopy.SpecialCells(xlCellTypeLastCell).Address
	
	wkbk.Worksheets(paste_wksht).Activate
	set veryLastCellPaste = wkbk.Worksheets(paste_wksht).UsedRange
	
	LastColLetterPaste = split(veryLastCellPaste.SpecialCells(xlCellTypeLastCell).Address(1,0), "$")(0)
	LastRowNumPaste = split(veryLastCellPaste.SpecialCells(xlCellTypeLastCell).Address(1,0), "$")(1)
	
	if InStr(UCase(paste_loc), "V") > 0 then 
		wkbk.Worksheets(copy_wksht).Range("A1:" & LastCellCopyAddress).copy wkbk.Worksheets(paste_wksht).Range("A" & LastRowNumPaste).offset(vert_offset,horz_offset) ' For vertical stacking 
	Elseif InStr(UCase(paste_loc), "H") > 0 then 
		wkbk.Worksheets(copy_wksht).Range("A1:" & LastCellCopyAddress).copy wkbk.Worksheets(paste_wksht).Range(LastColLetterPaste & "1").offset(vert_offset,horz_offset)  ' For horizontal stacking
	end if
	
	wkbk.Worksheets(paste_wksht).Cells.EntireColumn.Autofit
	eapp.CutCopyMode = False
	

	if InStr(UCase(del_sheet), "YES") > 0 then
		wkbk.Worksheets(copy_wksht).delete
	end if
	
	wkbk.Worksheets(1).Activate
	wkbk.Save

eapp.Workbooks.Open(Input_Excel).Close
eapp.Quit
set wkbk = nothing