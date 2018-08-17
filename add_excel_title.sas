/*
PURPOSE: 
		Create a merged and centered title on the first table of a sheet. 

INPUTS:
		excel_filename 		= Excel file name that has the tabs to be stacked. 
		excel_location		= Folder location of excel file.
		filetype			= (Default = xlsb) 
		sheet_name  		= Name of the sheet with table to be titled 
		sheet_title			= Desired title of the table. 
*/

%macro add_excel_title(excel_filename = , 
						excel_location = ,
						filetype = xlsb,
						sheet_name = ,
						sheet_title = );

    %macro dummy ;%mend dummy;
	%if %sysfunc(fileexist(&excel_location.\&excel_filename..&filetype)) %then
		X "'add_excel_title.vbs' 
		""&excel_location.\&excel_filename..&filetype"" ""&sheet_name"" ""&sheet_title"" ";
	%else %put %str(E)RROR: The excel file does not exist!;
%mend;
