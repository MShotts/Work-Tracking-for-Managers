
%let charyear=J;	/*The letter designation of the current year*/
%let recform=51;	/*The most recent form analyzed*/


%let QCdate=%sysfunc(today());  %let QCdate=%sysfunc(putn(&QCdate,yymmdd7.));
libname Database odbc noprompt="uid=Matthew; pwd=xxxxxxxx;dsn=MyDB;" schema=dbo stringdates=yes;
options mprint symbolgen obs=max;


%macro source(let,num);
data &let.Year;
do i=1 to &num.;
FNum=i;
length SAForm $4;
SAForm=compress("&let."||FNum,' ');
drop i FNum;
output;
end;
run;

%mend source;
/*Form year designation and then the number of forms for that year*/
%source(C,41);
%source(D,42);
%source(E,42);
%source(F,44);
%source(G,45);
%source(H,59);
%source(I,60);
%source(J,61);

data Source;
set CYear DYear EYear FYear GYear HYear IYear JYear;
length SMarker $6.;
SMarker='Source';
run;


/*The following extracts the available forms in MyDB*/
%macro cntry(c);
proc sql;
	create table &c.MyDBforms as
	select form
	,substr(form,2,1)||substr(form,5) as TAForm
/*	,'MyDB' as TMarker format=$6.*/
	,count(form) as &c.Vol 
	from Database.TestTakers
	where country_code="&c."
	group by Form
	order by Form;
quit;
%mend cntry;
%cntry(J);
%cntry(K);

data KTForms;
merge KMyDBForms JMyDBForms;
by TAForm;
if KVol ne . and JVol ne . then TAForm=compress(TAForm||"K",' ');
if KVol=. then delete;
drop JVol;
run;

proc datasets lib=work nolist; modify KTForms; rename KVol=MyDBVol; run; quit;
proc datasets lib=work nolist; modify JMyDBForms; rename JVol=MyDBVol; run; quit;

data MyDBforms; set JMyDBForms KTForms; run;


/*The following gathers all the names of the feedback files*/
filename DIRLIST1 pipe 'dir "B:\Genasys\Feedback\*" ';     
                                                             
data FdBck;                                               
infile dirlist1 lrecl=200 truncover;                          
input line $200.;                                            
if input(substr(line,1,10), ?? mmddyy10.)=. then delete;
if scan(line,2,'.') not in  ("out","cls") then delete;
length file_name $40 FAForm $12.;                                      
file_name=scan(line,-1," ");
FAForm=scan(upcase(file_name),1,"FE");
keep file_name FAForm;
run;


/*Read in the PIN file*/
proc import out=PINs
datafile='B:\Documentation\History\Pinned Items from May06.xls'
dbms=excel replace;
sheet="Pinned Items From May 2006";
mixed=yes; 
run;

proc sql;
	create table PINs2 as
	select substr(F2,2,1)||substr(F2,5) as PAForm
			,F3 as LC_PIN
			,F4 as RC_PIN
		from PINs
		where upcase(substr(F2,2,1)) not in ("A","B")
	order by PAForm
;
quit;


/*Read in the IAHist to see if updated*/
libname SASHist "B:\Documentation\history\SAS Backup";
proc sql;
	create table IAHist as
	select substr(Form,2,1)||substr(Form,5) as IAForm
	,count(Form) as IAVol
	,case when calculated IAVol=5 then "Hist OK"
	else "Hist XX"
	end as HistQC 
	from SASHist.IAHist
	where Form ne ""
	group by calculated IAForm
	order by calculated IAForm;
quit;

/*Read in the Scores History to see if updated*/
proc sql;
	create table ScoreHist as
	select substr(Form,2,1)||substr(Form,5) as ScForm
	,count(Form) as ScVol
	from SASHist.History
	where substr(Form,2,1) not in ("A","B")
	group by calculated ScForm
	order by calculated ScForm;
quit;

data ScoreHist2;
set ScoreHist;
where ScVol=2;
ScForm=compress(ScForm||"K",' ');
run;

data ScoreHist3; set ScoreHist ScoreHist2; run;


/*Merge all the results together*/
proc sql sortseq=linguistic;
	create table All3 as
	select a.*
	,b.FAForm
	,case when substr(A.SAForm,1,1)="C" then "C Year archived"
	else B.file_name
	end as Fdbk_File
	,c.*
	,d.*
	,e.IAForm
	,e.HistQC
	,f.*

	,case when SAForm='' then "Form not listed in Source"
	else "OK"
	end as Source_QC
	,case when calculated Fdbk_File='' then "No feedback file"
	else calculated Fdbk_File
	end as Fdbk_Status
	,case when C.MyDBVol=. then "Not in MyDB"
	else put(C.MyDBVol,12.)
	end as MyDB_Status
	,case when d.LC_PIN='' or d.RC_PIN='' then "No PIN info"
	else "OK"
	end as PIN_Status
	,case when e.HistQC='Hist XX' then "IA Hist does not have 5 stats"
	when e.HistQC='' then "IA Hist Missing"
	else "OK"
	end as IAH_Status
	,case when f.ScVol=. then "Score History Missing"
	else "OK"
	end as ScrH_Status
	
		from Source as A
		full outer join FdBck as B on A.SAForm=B.FAForm
		full outer join MyDBforms as C on A.SAForm=C.TAForm
		full outer join PINs2 as D on A.SAForm=D.PAForm
		full outer join IAHist as E on A.SAForm=E.IAForm
		full outer join ScoreHist3 as F on A.SAForm=F.ScForm
/*	Removing C Year since primary concern is 2012 and on*/
	where substr(SAForm,1,1) ne "C"
	order by A.SAForm
;
quit;


data Incomp Comp;
set All3;
if Source_QC ne "OK" or 
		Fdbk_Status = "No feedback file" or 
		MyDB_Status = "Not in MyDB" or 
		PIN_Status = "No PIN info" or 
		IAH_Status ne "OK" or
		ScrH_Status ne "OK" then output Incomp;
	else output Comp;
run;

data Incomp2;
set Incomp;
if substr(SAForm,1,1)="&charyear." and input(scan(substr(SAForm,2),1,"K"),4.)>&recform. then delete;
run;



ods listing close;
goptions reset=all;

/*Exporting results to Excel*/
ODS EXCEL FILE="B:\Users\Matthew\Post Admin Status_&QCdate..xlsx" style=sasdocprinter
options(sheet_name="Pending"
embedded_titles='yes');

Proc report data=Incomp2 nowd headskip split='~' style(header)={bordercolor=black background=aliceblue cellheight=0.45in};
Columns SAForm 
Fdbk_Status MyDB_Status PIN_Status IAH_Status ScrH_Status;
Define SAForm/style(column)={just=center};
Compute Fdbk_Status;
 		if index(Fdbk_Status,"No") ge 1 then call define("Fdbk_Status", "style", "style=[backgroundcolor=Orange]");
endcomp;
Compute MyDB_Status;
 		if index(MyDB_Status,"Not") ge 1 then call define("MyDB_Status", "style", "style=[backgroundcolor=Orange]");
endcomp;
Compute PIN_Status;
 		if index(PIN_Status,"No") ge 1 then call define("PIN_Status", "style", "style=[backgroundcolor=Orange]");
endcomp;
Compute IAH_Status;
 		if IAH_Status ne "OK" then call define("IAH_Status", "style", "style=[backgroundcolor=Orange]");
endcomp;
Compute ScrH_Status;
 		if ScrH_Status ne "OK" then call define("ScrH_Status", "style", "style=[backgroundcolor=Orange]");
endcomp;
run;

ODS EXCEL options(sheet_name="Complete");

Proc report data=Comp nowd headskip split='~' style(header)={bordercolor=black background=aliceblue cellheight=0.45in};
Columns SAForm Fdbk_Status MyDB_Status PIN_Status IAH_Status ScrH_Status;
Define SAForm/style(column)={just=center};
Define Fdbk_Status/style(column)={just=center};
Define MyDB_Status/style(column)={just=center};
Define PIN_Status/style(column)={just=center};
Define IAH_Status/style(column)={just=center};
Define ScrH_Status/style(column)={just=center};
run;

ods excel close;
