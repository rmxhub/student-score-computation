/***************************************************************************************

The class Project of STAT 5533

The trick part in this project is how to select out the highest 15 out of 18 to 20 times of HW and
quizzes in each class. In other words, a student can drop 3 to 5 homework assignments or
quizzes depending on the class s/he is with. Here I provide Four ways to calculate the
total scores of homework and quizzes.

Written By Rongmin Xia  
***************************************************************************************/

option ls=120 nodate;

%let dir="\\math\courses\Statistics\STAT5533\project_one.xlsx";
* dir="C:\Users\Administrator\Desktop\UHCL\2014_9\5533_sas\project\project_one.xlsx" ;


%let outdata=student4;  
%let sheetname="class1";
%let outdata_tmp=studentt;

/*
Rename the column name with new one, here is not for one name, but for entire sheet.
*/
%macro rename_RX(oldvarlist, newvarlist);
  %let k=1;
  %let old = %scan(&oldvarlist, &k);
  %let new = %scan(&newvarlist, &k);
     %do %while(("&old" NE "") & ("&new" NE ""));
      rename &old = &new;
      label &old = &new;
    %let k = %eval(&k + 1);
      %let old = %scan(&oldvarlist, &k);
      %let new = %scan(&newvarlist, &k);
%put &old &new;
%end;
%mend rename_RX;

/*
Readandprocess(dirfilename, sheetname,outdata, classno, method): Read one sheet from Excel 
dirfilename: file name and directory
outdata: the output dataset read from Excel
sheetname: the sheet name inside Excel
classno: assign the class number for each sheet
Method:   sort method, value range 1,2,3,4
outdata_tmp: temperary dataset used during compution*/

%Macro readandprocess(dirfilename, sheetname,outdata, classno, method);
/* 
  Read one sheet from Excel 
*/
PROC IMPORT OUT=&outdata_tmp
      DATAFILE= &dirfilename
      DBMS=xlsx REPLACE;
      sheet=&sheetname;
      getnames=yes;
RUN;
/*
Get the table information, such as column number, observation number, list of column name.
(aslo you can use SQL with dictionary or Proc Contents)
*/
%let dsid  = %sysfunc(open(&outdata_tmp,i));
%let nvars = %sysfunc(attrn(&dsid,NVARS));
%let nobs  = %sysfunc(attrn(&dsid,NOBS));
%let varlist=;
%do i=1 %to &nvars;
    %let varlist=&varlist %sysfunc(varname(&dsid,&i));
%end;

%let rc = %sysfunc(close(&dsid));  
%let nonHWn=4;
%put varlist=&varlist;
%put nvars=&nvars;
/*
rename the column name and make it as: Student_ID Test_1 Test_2 Final HW1-HW...
*/
Data class_2;
  set &outdata_tmp;
  %let HWrange=;
  %do i=1 %to 30;
     %let HWrange=&HWrange HW&i;
  %end;
    %let newlist=Student_ID Test_1 Test_2 Final &HWrange;
    %rename_RX(&varlist, &newlist);
  %put HWrange=&HWrange;
run;

/*
Four kind of sort method. 
The 1st is based on Bubble sort, comparing each pair of adjacent items and swapping them if they are in the wrong order.
The 2nd is based on array tranformation with Proc sort.
The 3rd is based on the Proc transpose in multi-columns.
  Using proc transpose to switch rows to columns, then sort each column, and combine all sorted column to a new dataset
The 4th is based on Sortn function, which can sort in row direction.
The 5th which is done in Project one, is as transposing to two column (student_ID and HWScore), then sort the HWScore, and transpose it back. 

*/
%if &method=1 %then %do;/*Bubble sort*/
   data class_new_2;
      set class_2;
      array S{*} HW:;
      do i=1 to dim(S)-1;
          do j=i+1 to dim(S);
            if S{i} < S{j} then do;
                temp=S{i};
                S{i}=S{j};
                S{j}=temp;
            end;
          end;
      end;
      drop i j temp;
    run;
  %end;
%else %if &method=2 %then %do; /*array tranformation with Proc sort*/
  DATA class_new_2;             /*HW columns transform : multi-single, keep the Student_ID column*/
    SET class_2;
    array S{*} HW:;
    DO i = 1 TO dim(S);
      SCORE = S[i];
      OUTPUT;
    END;
    KEEP Student_ID SCORE;
  RUN;

  Proc sort data=class_new_2 out =class_new_2_sort; /* Sort the single column HW by Student_ID*/
    by Student_ID descending SCORE;

  DATA class_new_2;  /*HW columns transform : single-multiple, the key is that same Student_ID will be output as one Row  */
      nHW = &nvars-&nonHWn;
      RETAIN HW1-HW30;
      array S{*} HW:;
      SET class_new_2_sort;
      BY Student_ID;
      i2=_N_ - floor(_N_ / nHW)*nHW;
      if i2=0 then i2=nHW;
        S{i2} = SCORE;
      IF LAST.Student_ID THEN OUTPUT;
      KEEP Student_ID HW1-HW15;
    RUN;

    proc datasets library=work; /*delete the temp dataset*/
      delete class_new_2_sort;
    run;
  %end;
  %else %if &method=3 %then %do; /*good*/
    proc transpose data=class_2(keep=HW:) out=class_new_2; run; /*transpose to switch rows to columns*/

    %do i=1 %to &nobs;
      proc sort data=class_new_2(keep=Col&i) out=column;  /*sort one column and combine all sorted columns*/
        by descending Col&i;
      run;
      data sortedrows;                  
       %if &i>1 %then set  sortedrows;;
       set  column;
      run;
    %end;

    proc transpose data=sortedrows out=class_new_2; run; /*transpose to switch columns to rows */

    Data class_new_2;
      set class_2(keep=Student_ID Test_1 Test_2 Final);
      set class_new_2(drop=_:);
      rename Col1-Col15 = HW1-HW15;
    run;
    proc datasets library=work;
      delete sortedrows;
      delete column;
    run;

  %end;
  %else %if &method=4 %then %do; /*Sortn function to sort in ascending direction, and then change to descending direction using array*/
    Data class_new_2;
      set class_2;
      array S{*} HW:;
      call sortn(of S[*]);
      do i=1 to floor(DIM(s)/2);
        tmp = S[i];
        S[i] = S[DIM(s)-i+1];
        S[DIM(s)-i+1]=tmp;
       end;
      drop i tmp;
      call symputx('array_x',catx(' ',of S[*]));
    run;
    %put &=array_x;
  %end;

  /*Compute the Grade*/
  data &outdata;
    retain Student_ID Grade Grade_value  Test_1 Test_2 Final HW_Total;
    set class_2(keep = Student_ID Test_1 Test_2 Final);
    set class_new_2(keep = HW1-HW15);
    HW_Total=sum(of HW1 - HW15);
    Grade_value=HW_Total*0.2+Test_1*.2+Test_2*.2+Final*.2;

    if Grade_value >90 then do;
       Grade ='A'; end;
    else if Grade_value >80 then do;
       Grade ='B'; end; 
    else if Grade_value >70 then do;
        Grade ='C'; end;  
    else if Grade_value >60 then do;
        Grade ='D'; end; 
    else do;
        Grade ='F'; end; 
    classNo=&classno;
  run;

  proc datasets library=work;
    delete class_2;
    delete class_new_2;
  run;/**/


%Mend readandprocess;

/*
Readandprocess(dirfilename, sheetname,outdata, classno, method)  
dirfilename: file name and directory
outdata: the output dataset read from Excel
sheetname: the sheet name inside Excel
classno: assign the class number for each sheet
Method:   sort method, value range 1,2,3,4
outdata_tmp: temperary dataset used during compution*/

%readandprocess(&dir,"class1",student1,1,1);
proc print; run;

%readandprocess(&dir,"class2",student2,2,2);
proc print; run;

%readandprocess(&dir,"class3",student3,3,3);
proc print; run;

*Combine all data;
Data student_w;
  set student1-student3; 
run;

proc print; run;

Proc Means data=student_w N MAX MIN MEAN STD;
  var Grade_value ;
run;

Proc TTest data=student_w  h0=80 alpha=0.05 ;  
  var Grade_value;
  by classNo;
run;


Proc Univariate data=student_w  normal plot;
  var Grade_value3;
  QQPLOT Grade_value3 /NORMAL(MU=EST SIGMA =EST COLOR =RED ); 
  HISTOGRAM /normal; 
run;

proc glm   data=student_w;
   class classNo;
   model Grade_value3 = classNo;
   means classNo / Tukey  alpha = 0.05 HOVTEST;
run;

axis1  label=(a=90 'Grade');
axis2 label=( 'Frequency');                                                                                   
proc gchart data= student_w;
  hBAR  grade /maxis=axis1 raxis=axis2;
run;
proc freq data=student_w;
  table  Grade;
run;

proc reg data=student_w;
  model Grade_value= test_1 Test_2 Final HW_Total;
run;

proc corr data=student_w;
  var Grade_value;
  with test_1 Test_2 Final HW_Total;
run;

Quit;
