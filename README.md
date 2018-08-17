# utility-macros
Tedious tasks made easier by automation. I wrote these macros while working at [Equity Methods](https://www.equitymethods.com/) to help with the final formatting of reports that we deliver to clients, and to inrease productivity and reduce the risk of human error by automating manual and tedious tasks. These macros have saved me and my colleagues many hours of work, and hopefully you will find them as useful as we have. 

## Summary of macros and their functions
<b><i>%create_generic_labels</i></b>: When exporting reports, the variable's label can be used as the header of the column in Excel. This macro takes the variable's name and converts the underscores to spaces, propcases each word, and assigns this string to the label of the variable. This quickly prepares a dataset for export with nice looking headers, and saves the work of manually assigning a label to each variable. (Example: a variable named the_answer_to_life_universe_everything would receive the label "The Answer To Life Universe Everything".)

<b><i>%stack_excel_sheets</i></b>: This calls a vbscript that copies and pastes the useful contents (not the whitespace) from one Excel sheet into another. This allows multiple tables to be displayed in a single tab, instead of having separate tab for each exported dataset (or having to manually combine tabs by hand, as is usually the case), and is particularly useful for tabs that display summary tables. The macro includes the option to delete the tab copied (default), and whether to paste the copied tab vertically or horizontally (and the offset, if desired). As an example, you can copy the contents of "Sheet 2" and paste it underneath (or to the right of) the contents of "Sheet 1", and then delete "Sheet 2". 

<b><i>%add_excel_title</i></b>: This calls a vbscript that creates a merged and centered cell that acts as the title of a table. This was written to complement <b><i>%stack_excel_sheets</i></b> because it adds clarity to each table if there are multiple tables on a single tab.

<b><i>%bold_excel_keywords</i></b>: This calls a vbscript that finds keywords and bolds their respective rows. An example might be a row that contains the word "Subtotal".

### Prerequisites
For <b><i>%stack_excel_sheets</i></b>, <b><i>%add_excel_title</i></b>, and <b><i>%bold_excel_keywords</i></b>, you must update the filepaths in the SAS code to the location of the vbscript file. (example: Code currently reads: 'stack_excel_sheets.vbs', and needs to be updated to wherever the vbscript file is stored, so it could be updated to 'C:\SAS Macros\stack_excel_sheets.vbs') 

For <b><i>%create_generic_labels</i></b>, this code utilizes Ted Clay's array macro, and must be included for the code to work. See http://www2.sas.com/proceedings/sugi31/040-31.pdf for more information about the macro, and see https://gist.github.com/JoostImpink/c22197c93ecd27bbf7ef for a copy of the code. 
