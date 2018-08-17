/*
PURPOSE: 
		Bold entire rows based on keywords (such as "subtotal"). This will search the entire sheet for keywords.

INPUTS:
		excel_filename 		= Excel file name 
		excel_location		= Folder location of excel file.
		filetype			= (Default = xlsb) 
		sheet_name  		= Name of the sheet search
		keywords			= Keywords to search and have rows bolded, separated with pipes | 
								Note: Requires an exact match of the cell contents.
TO DO:
		1. The keyword loop would be much more efficient in the vbscript file instead of SAS.
*/

%macro bold_excel_keywords(excel_filename = , 
						excel_location = ,
						filetype = xlsb,
						sheet_name = ,
						keywords = );

    %macro dummy ;%mend dummy;
	%if %sysfunc(fileexist(&excel_location.\&excel_filename..&filetype)) %then %do;
		%let num_keywords = %sysfunc(countw(&keywords,%str(|),Q));

		%do n_kw = 1 %to &num_keywords;
			%let kw = %scan(&keywords, &n_kw, %str(|));
			X "'bold_excel_keywords.vbs' 
			""&excel_location.\&excel_filename..&filetype"" ""&sheet_name"" ""&kw"" ";
		%end;
	%end;
	%else %put %str(E)RROR: The excel file does not exist!;
%mend;