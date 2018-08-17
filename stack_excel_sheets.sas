/*
PURPOSE: 
		To copy the entire useful contents (not the blanks to the edges of the sheet)
		of a desired tab and paste them into other another designated tab.

INPUTS:
		excel_filename 			= Excel file name that has the tabs to be stacked. 
		excel_location			= Folder location of excel file.
		filetype				= (Default = xlsb) 
		sheet_to_copy 			= Tab on excel file that is to be copied. 
		sheet_to_paste 			= Tab on excel file that will receive the copied cells from tab_to_copy. 
		paste_position 			= (DEFAULT = V) Choose either V (for Vertical) or H (for horizontal) (not case sensitive) 
									for paste location of copied cells. 
		delete_copied_sheet 	= (DEFAULT = Yes) Choose Yes (not case sensitive) to delete the tab that contains the copied material.
		vertical_offset			= Number of rows below given location to paste (default of 1 for vertical, 0 for horizontal)
		horizontal_offset 		= Number of columns to right of given location to paste (default of 0 for vertical, 1 for horizontal)
TO DO:
		1. Could be useful to be able to copy/paste tabs from one workbook into another. 
		2. Turn all filters off of pasted sheet. 
*/

%macro stack_excel_sheets(excel_filename = , 
						excel_location = ,
						filetype = xlsb,
						sheet_to_copy = , 
						sheet_to_paste = , 
						paste_position = V, 
						delete_copied_sheet = Yes,
						vertical_offset = ,
						horizontal_offset = );

    %macro dummy ;%mend dummy;
	%local vertical_offset horizontal_offset;

	%*Create default offset positions unless specified;
	%if &paste_position = V %then %do;
		%if &vertical_offset = %then %let vertical_offset = 1;
		%if &horizontal_offset = %then %let horizontal_offset = 0;
	%end;
	%if &paste_position = H %then %do;
		%if &vertical_offset = %then %let vertical_offset = 0;
		%if &horizontal_offset = %then %let horizontal_offset = 1;
	%end;

	%if %sysfunc(fileexist(&excel_location.\&excel_filename..&filetype)) %then
		X "'stack_excel_sheets.vbs' 
		""&excel_location.\&excel_filename..&filetype"" ""&sheet_to_copy"" ""&sheet_to_paste"" ""&paste_position"" ""&delete_copied_sheet"" 
			""&vertical_offset"" ""&horizontal_offset"" ";
	%else %put %str(E)RROR: The excel file does not exist!;
%mend;