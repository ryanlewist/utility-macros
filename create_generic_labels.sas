/*
PURPOSE
		Quickly and easily create labels for variables by replacing underscores 
		with spaces, and proper casing the label. 

INPUTS
		adj_all  	= 	Set to Y if you want to relabel all variables
						by using this method of replacing underscores with spaces and 
						proper casing the words. Otherwise, variables with existing labels 
						will not be adjusted.
		upcase_two  = 	Set this to Y if you would like to uppercase all words that contain 
						two or fewer letters. (This is useful for abbreviations such as ID, AZ, CA, NY, etc.
*/

%macro create_generic_labels(indsn= , outdsn = &indsn, adj_all = N, upcase_two = Y);
    %macro dummy ;%mend dummy;
	proc contents data=&indsn noprint out=_current_labels (keep=name label); run;
	data _new_labels (keep=name new_label);
		set _current_labels;
		if upcase("&adj_all") = "Y" then new_label = propcase(tranwrd(name,"_"," "));
		else if label = "" then do;
			new_label = propcase(tranwrd(name,"_"," "));
			if upcase("&upcase_two") = "Y" then do;
				word_num = 1;
				length upcase_two_label $100.;
				do until(word_num > countw(new_label));
					if length(scan(new_label,word_num)) le 2 then upcase_two_label = catx(" ",upcase_two_label, upcase(scan(new_label,word_num)));
					else upcase_two_label = catx(" ",upcase_two_label, scan(new_label,word_num));
					word_num + 1;
				end;
				new_label = upcase_two_label;
			end;
		end;
		else new_label = label;
	run;

	%array(array=name, data=_new_labels, var=name); %*This is Ted Clay's array macro. See http://www2.sas.com/proceedings/sugi31/040-31.pdf;
	%array(array=new_label, data=_new_labels, var=new_label); 

    %local openDS num_obs closeDS; 

	%let openDS = %sysfunc(open(_new_labels));
	%let num_obs  = %sysfunc(attrn(&openDS,nlobs));
	%let closeDS = %sysfunc(close(&openDS));

	data &outdsn;
		set &indsn;
		%do i = 1 %to &num_obs;
			label "&&name&i"n = "&&new_label&i";
		%end;
	run;

	proc datasets lib=work nolist;
		delete _new_labels _current_labels;
	run;

%mend;
