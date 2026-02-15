%let pgm=utl-altair-slc-determining-type-and-length-of-excel-columns-before-importing;

%stop_submission;

Altair slc determining type and length of excel columns before importing

too long to post here, see gothub
https://github.com/rogerjdeangelis/utl-altair-slc-determining-type-and-length-of-excel-columns-before-importing

CONTENTS

   1 create excel odbc dsn
   2 input sheet and named range
   3 query for type length
   4 create attrib statement
   5 use attrib to make table
   6 query named-range
   7 query using proc r
   8 drop down powershell


PROBLEM

  Use Microsoft Access SQL functions like 'isnumeric' and 'length' to get type and length for each column in a excel sheet or named range.
  Finally use type anf length to import the sheet or table.

                   INPUT                                  OUTPUT (CREATE TABLE CLASS with o
                   -----                                  ------

  d:/xls/class.xlsx                                   data workx.class;

  -------------------------+                             informat &attrib; /*--- NAM $7. SEX $1. AGE $2. WGT 3. ---*/
  | A1| fx       |  NAME   |                             format &attrib;   /*--- NAM $7. SEX $1. AGE $2. WGT 3. ---*/
  ---------------------------------------------
  [_] |    A     |    B    |    C    |   D    |          set xls.class;
  ---------------------------------------------
   1  | NAME     |   SEX   |   AGE   | HEIGHT |       run;
   -- |----------+--------+---------+---------+
   2  |^ Alfred  |^  0     |^  UU   |^  100   |
   -- |----------+---------+--------+---------+
   3  |^ Alice   |^  1     |^  13   |^  122   |
   -- |----------+---------+--------+---------+
   4  |^ Barbara |^  F     |^  13   |^  111   |
   -- |----------+---------+--------+---------+
   5  |^ Carol   |^  F     |^  14   |^  113   |
   -- |----------+---------+--------+---------+
   6  |^ Henry   |^  M     |^  14   |^  121   |
   -- |----------+---------+--------+---------+
   7  |^ James   |^  M     |^  12   |^  116   |
   -- |----------+---------+--------+---------+
   8  |^ Jane    |^  F     |^  12   |^  109   |
   -- |----------+--------+---------+---------+
  [CLASS]


SOAPBOX ON

This cost me hours.
My powershell script below worked in win 10 but fails in win 11.
You have to run the script as administrator.

STATEMENT FROM MICRISOFT

Permissions: Must run as Administrator;
Win11 enforces this more strictly for registry writes under HKCU\Software\ODBC\ODBC.INI.

Would be nice if MS offered an option to install Win 11 without reduced functionality.
Also I would like the macro recorder for powerpoint back.

SOAPBOX OFF

/*                       _                   _ _
/ |   ___ _ __ ___  __ _| |_ ___    ___   __| | |__   ___
| |  / __| `__/ _ \/ _` | __/ _ \  / _ \ / _` | `_ \ / __|
| | | (__| | |  __/ (_| | ||  __/ | (_) | (_| | |_) | (__
|_|  \___|_|  \___|\__,_|\__\___|  \___/ \__,_|_.__/ \___|

*/

/*--- This works with admin priviledges             ---*/
/*--- macro utl_submit_ps64 on end and in this repo ---*/

%utl_submit_ps64('
Add-OdbcDsn -Name "class" -DriverName "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)" -DsnType "User" -SetPropertyValue "Dbq=d:\xls\havw.xlsx"; 
 Get-OdbcDsn; ');

/*--- or open a command window as admin and paste powershell script. ---*/
/*--- Or open USER odbc 64bit and mouse serf.                       ---*/

/*
| | ___   __ _
| |/ _ \ / _` |
| | (_) | (_| |
|_|\___/ \__, |
         |___/
*/

1   Altair SLC      08:49 Tuesday, February 10, 2026

NOTE: Copyright 2002-2025 World Programming, an Altair Company
NOTE: Altair SLC 2026 (05.26.01.00.000758)
      Licensed to Roger DeAngelis
NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
NOTE: AUTOEXEC source line
1       +  ï»¿ods _all_ close;
           ^
ERROR: Expected a statement keyword : found "?"
NOTE: Library workx assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\wpswrkx

NOTE: Library slchelp assigned as follows:
      Engine:        WPD
      Physical Name: C:\Progra~1\Altair\SLC\2026\sashelp

NOTE: Library worksas assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\worksas

NOTE: Library workwpd assigned as follows:
      Engine:        WPD
      Physical Name: d:\workwpd


LOG:  8:49:50
NOTE: 1 record was written to file PRINT

NOTE: The data step took :
      real time : 0.031
      cpu time  : 0.015


NOTE: AUTOEXEC processing completed

1
2         %utl_submit_ps64('
3         Add-OdbcDsn `
              -Name "class" `
              -DriverName "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)" `
              -DsnType "User" `
              -SetPropertyValue "Dbq=d:\xls\have.xlsx"
4         Get-OdbcDsn;
5         ')

NOTE: The file py_pgm is:
      Filename='d:\wpswrk\_TD17768\py_pgm.ps1',
      Owner Name=BUILTIN\Administrators,
      File size (bytes)=0,
      Create Time=08:49:50 Feb 10 2026,
      Last Accessed=08:49:50 Feb 10 2026,
      Last Modified=08:49:50 Feb 10 2026,
      Lrecl=32766, Recfm=V

Add-OdbcDsn `
    -Name "class" `
    -DriverName "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)" `
    -DsnType "User" `
    -SetPropertyValue "Dbq=d:\xls\have.xlsx"
Get-OdbcDsn;


NOTE: 2 records were written to file py_pgm
      The minimum record length was 384
      The maximum record length was 384
NOTE: The data step took :
      real time : 0.000
      cpu time  : 0.015


d:\wpswrk\_TD17768\py_pgm.ps1

NOTE: The infile rut is:
      Unnamed Pipe Access Device,
      Process=powershell.exe -executionpolicy bypass -file d:\wpswrk\_TD17768\py_pgm.ps1 ,
      Lrecl=32767, Recfm=V


Name       : class
DsnType    : User
Platform   : 64-bit
DriverName : Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)
Attribute  : {DBQ, DriverId, ImplicitCommitSync, Threads...}

Name       : dBASE Files
DsnType    : User
Platform   : 64-bit
DriverName : Microsoft Access dBASE Driver (*.dbf, *.ndx, *.mdx)
Attribute  : {Threads, SafeTransactions, ImplicitCommitSync, DriverId...}

Name       : Excel Files
DsnType    : User
Platform   : 64-bit
DriverName : Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)
Attribute  : {SafeTransactions, DriverId, ImplicitCommitSync, Threads...}

Name       : MS Access Database
DsnType    : User
Platform   : 64-bit
DriverName : Microsoft Access Driver (*.mdb, *.accdb)
Attribute  : {Threads, SafeTransactions, ImplicitCommitSync, DriverId...}

Name       : class
DsnType    : User
Platform   : 64-bit
DriverName : Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)
Attribute  : {DBQ, DriverId, ImplicitCommitSync, Threads...}

Name       : sqlitedsn
DsnType    : System
Platform   : 64-bit
DriverName : GM-Software SQLite3 ODBC Driver
Attribute  : {EnableViews, Exclusive, UseTriggers, NoFollow...}


NOTE: 46 records were written to file PRINT

NOTE: 46 records were read from file rut
      The minimum record length was 0
      The maximum record length was 73
NOTE: The data step took :
      real time : 1.046
      cpu time  : 0.015



ERROR: Error printed on page 1

NOTE: Submitted statements took :
      real time : 1.157
      cpu time  : 0.093

/*___    _                   _         _               _                                    _
|___ \  (_)_ __  _ __  _   _| |_   ___| |__   ___  ___| |_   _ __   __ _ _ __ ___   ___  __| |  _ __ __ _ _ __   __ _  ___
  __) | | | `_ \| `_ \| | | | __| / __| `_ \ / _ \/ _ \ __| | `_ \ / _` | `_ ` _ \ / _ \/ _` | | `__/ _` | `_ \ / _` |/ _ \
 / __/  | | | | | |_) | |_| | |_  \__ \ | | |  __/  __/ |_  | | | | (_| | | | | | |  __/ (_| | | | | (_| | | | | (_| |  __/
|_____| |_|_| |_| .__/ \__,_|\__| |___/_| |_|\___|\___|\__| |_| |_|\__,_|_| |_| |_|\___|\__,_| |_|  \__,_|_| |_|\__, |\___|
                |_|                                                                                             |___/
*/

%utlfkil(d:/xls/have.xlsx); * delete if exist - it works with an existing workbook;

options set=RHOME "C:\Progra~1\R\R-4.5.2\bin\r";
proc r;
submit;
library(openxlsx);
class <- read.table(header = TRUE,sep=",", text = "
NAM, SEX, AGE, WGT
Alfred ,0,UU,100
Alice ,1,13,122
Barbara,F,13,111
Carol ,F,14,113
Henry ,M,14,121
James ,M,12,116
Jane ,F,12,109
");
class
wb <- createWorkbook("d:/xls/have.xlsx")
addWorksheet(wb, "class")
writeData(wb, sheet = 1, x = class, startCol = 1, startRow = 1)
createNamedRegion(
  wb = wb,
  sheet = 1,
  name = "class",
  rows = 1:(nrow(class) + 1),
  cols = 1:ncol(class)
)
saveWorkbook(wb,"d:/xls/have.xlsx", overwrite = TRUE)
endsubmit;
run;


/*********************************************************************************************/
/* d:/xls/class.xlsx                                                                         */
/*                                                                                           */
/* NAMED-RANGE   Formulas -> Name Manager                                                    */
/*                                                                                           */
/*  NAME   VALUE                                                  REFERS TO         SCOPE    */
/*                                                                                           */
/*  CLASS  {"NAM","SEX","AGE","WGT";"Alfred","0","UU","100"...   =class!$A$1:$D$8   WORKBOOK */
/*                                                                                           */
/*-------------------------------------------------------------------------------------------*/
/*                                                                                           */
/* d:/xls/class.xlsx                                                                         */
/*                                                                                           */
/* -------------------------+                                                                */
/* | A1| fx       |  NAME   |                                                                */
/* ---------------------------------------------                                             */
/* [_] |    A     |    B    |    C    |   D    |                                             */
/* ---------------------------------------------                                             */
/*  1  | NAME     |   SEX   |   AGE   | HEIGHT |                                             */
/*  -- |----------+--------+---------+---------+                                             */
/*  2  |^ Alfred  |^  0     |^  UU   |^  100   |                                             */
/*  -- |----------+---------+--------+---------+                                             */
/*  3  |^ Alice   |^  1     |^  13   |^  122   |                                             */
/*  -- |----------+---------+--------+---------+                                             */
/*  4  |^ Barbara |^  F     |^  13   |^  111   |                                             */
/*  -- |----------+---------+--------+---------+                                             */
/*  5  |^ Carol   |^  F     |^  14   |^  113   |                                             */
/*  -- |----------+---------+--------+---------+                                             */
/*  6  |^ Henry   |^  M     |^  14   |^  121   |                                             */
/*  -- |----------+---------+--------+---------+                                             */
/*  7  |^ James   |^  M     |^  12   |^  116   |                                             */
/*  -- |----------+---------+--------+---------+                                             */
/*  8  |^ Jane    |^  F     |^  12   |^  109   |                                             */
/*  -- |----------+--------+---------+---------+                                             */
/* [CLASS]                                                                                   */
/*********************************************************************************************/

/*
| | ___   __ _
| |/ _ \ / _` |
| | (_) | (_| |
|_|\___/ \__, |
         |___/
*/

1                                          Altair SLC      12:25 Tuesday, February 10, 2026

NOTE: Copyright 2002-2025 World Programming, an Altair Company
NOTE: Altair SLC 2026 (05.26.01.00.000758)
      Licensed to Roger DeAngelis
NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
NOTE: AUTOEXEC source line
1       +  ï»¿ods _all_ close;
           ^
ERROR: Expected a statement keyword : found "?"
NOTE: Library workx assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\wpswrkx

NOTE: Library slchelp assigned as follows:
      Engine:        WPD
      Physical Name: C:\Progra~1\Altair\SLC\2026\sashelp

NOTE: Library worksas assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\worksas

NOTE: Library workwpd assigned as follows:
      Engine:        WPD
      Physical Name: d:\workwpd


LOG:  12:25:25
NOTE: 1 record was written to file PRINT

NOTE: The data step took :
      real time : 0.031
      cpu time  : 0.000


NOTE: AUTOEXEC processing completed

1         %utlfkil(d:/xls/have.xlsx); * delete if exist - it works with an existing workbook;
2
3         options set=RHOME "C:\Progra~1\R\R-4.5.2\bin\r";
4         proc r;
5         submit;
6         library(openxlsx);
7         class <- read.table(header = TRUE,sep=",", text = "
8         NAM, SEX, AGE, WGT
9         Alfred ,0,UU,100
10        Alice ,1,13,122
11        Barbara,F,13,111
12        Carol ,F,14,113
13        Henry ,M,14,121
14        James ,M,12,116
15        Jane ,F,12,109
16        ");
17        class
18        wb <- createWorkbook("d:/xls/have.xlsx")
19        addWorksheet(wb, "class")
20        writeData(wb, sheet = 1, x = class, startCol = 1, startRow = 1)
21        createNamedRegion(
22          wb = wb,
23          sheet = 1,
24          name = "class",
25          rows = 1:(nrow(class) + 1),
26          cols = 1:ncol(class)

2

27        )
28        saveWorkbook(wb,"d:/xls/have.xlsx", overwrite = TRUE)
29        endsubmit;
NOTE: Using R version 4.5.2 (2025-10-31 ucrt) from C:\Program Files\R\R-4.5.2

NOTE: Submitting statements to R:

> library(openxlsx);
> class <- read.table(header = TRUE,sep=",", text = "
+ NAM, SEX, AGE, WGT
+ Alfred ,0,UU,100
+ Alice ,1,13,122
+ Barbara,F,13,111
+ Carol ,F,14,113
+ Henry ,M,14,121
+ James ,M,12,116
+ Jane ,F,12,109
+ ");
> class
> wb <- createWorkbook("d:/xls/have.xlsx")
> addWorksheet(wb, "class")
> writeData(wb, sheet = 1, x = class, startCol = 1, startRow = 1)
> createNamedRegion(
+   wb = wb,
+   sheet = 1,
+   name = "class",
+   rows = 1:(nrow(class) + 1),
+   cols = 1:ncol(class)
+ )
> saveWorkbook(wb,"d:/xls/have.xlsx", overwrite = TRUE)

NOTE: Processing of R statements complete

30        run;
NOTE: Procedure r step took :
      real time : 2.270
      cpu time  : 0.031


ERROR: Error printed on page 1

NOTE: Submitted statements took :
      real time : 2.594
      cpu time  : 0.140


/*____                                 _                     _                  _   _
|___ /    __ _ _   _  ___ _ __ _   _  | |_ _   _ _ __   ___ | | ___ _ __   __ _| |_| |__
  |_ \   / _` | | | |/ _ \ `__| | | | | __| | | | `_ \ / _ \| |/ _ \ `_ \ / _` | __| `_ \
 ___) | | (_| | |_| |  __/ |  | |_| | | |_| |_| | |_) |  __/| |  __/ | | | (_| | |_| | | |
|____/   \__, |\__,_|\___|_|   \__, |  \__|\__, | .__/ \___||_|\___|_| |_|\__, |\__|_| |_|
            |_|                |___/       |___/|_|                       |___/
*/

proc sql;
  connect to odbc (dsn="class");

  create
     table workx.want as
  select
     *
  from connection to odbc (
    select
      count(*)                       as rows
     ,sum(iif(isnumeric(NAM),1, 0))  as num_nam
     ,sum(iif(isnumeric(SEX),1, 0))  as num_sex
     ,sum(iif(isnumeric(AGE),1, 0))  as num_age
     ,sum(iif(isnumeric(WGT),1, 0))  as num_wgt
     ,max(len(nam))                  as len_nam
     ,max(len(sex))                  as len_sex
     ,max(len(age))                  as len_age
     ,max(len(wgt))                  as len_wgt
    from `class$`
  );

  disconnect from odbc;
quit;

proc print data=workx.want;
run;quit;

/****************************************************************************************************/
/*  WORKX.TYPLEN total obs=1                                                                        */
/*                                                                                                  */
/*  ROWS   NUM_NAM   NUM_SEX   NUM_AGE   NUM_WGT      LEN_NAM      LEN_SEX      LEN_AGE     LEN_WGT */
/*  ----------------------------------------------------------------------------------------------- */
/*     7         0         2         6         7            7            1            2           3 */
/*                                                                                                  */
/*            HOW NAMY OF THE ROWS ARE NUMERIC               WHAT IS THE LONGEST LENGTHS            */
/*                                                                                                  */
/*            CHAR      CHAR       CHAR       NYM        7 bytes    1 byte      2 bytes  3 num fmt  */
/*          ----------------------------------------     ----------------------------------------   */
/*  ROWS    NUM_NAM    NUM_SEX    NUM_AGE    NUM_WGT     LEN_NAM    LEN_SEX    LEN_AGE    LEN_WGT   */
/*                                                                                                  */
/*    7        0          2          6          7           7          1          2          3      */
/*  IF <7 USE CHARACTER FORMAT                                                                      */
/****************************************************************************************************/

/*
| | ___   __ _
| |/ _ \ / _` |
| | (_) | (_| |
|_|\___/ \__, |
         |___/
*/

1                                          Altair SLC      12:27 Tuesday, February 10, 2026

NOTE: Copyright 2002-2025 World Programming, an Altair Company
NOTE: Altair SLC 2026 (05.26.01.00.000758)
      Licensed to Roger DeAngelis
NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
NOTE: AUTOEXEC source line
1       +  ï»¿ods _all_ close;
           ^
ERROR: Expected a statement keyword : found "?"
NOTE: Library workx assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\wpswrkx

NOTE: Library slchelp assigned as follows:
      Engine:        WPD
      Physical Name: C:\Progra~1\Altair\SLC\2026\sashelp

NOTE: Library worksas assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\worksas

NOTE: Library workwpd assigned as follows:
      Engine:        WPD
      Physical Name: d:\workwpd


LOG:  12:27:10
NOTE: 1 record was written to file PRINT

NOTE: The data step took :
      real time : 0.031
      cpu time  : 0.015


NOTE: AUTOEXEC processing completed

1         proc sql;
2           connect to odbc (dsn="class");
NOTE: Connected to DB: class (EXCEL version 12.00.0000)
NOTE: Connected to DB: class (EXCEL version 12.00.0000)
NOTE: Successfully connected to database ODBC as alias ODBC.
3
4           create
5              table workx.want as
6           select
7              *
8           from connection to odbc (
9             select
10              count(*)                       as rows
11             ,sum(iif(isnumeric(NAM),1, 0))  as num_nam
12             ,sum(iif(isnumeric(SEX),1, 0))  as num_sex
13             ,sum(iif(isnumeric(AGE),1, 0))  as num_age
14             ,sum(iif(isnumeric(WGT),1, 0))  as num_wgt
15             ,max(len(nam))                  as len_nam
16             ,max(len(sex))                  as len_sex
17             ,max(len(age))                  as len_age
18             ,max(len(wgt))                  as len_wgt
19            from `class$`
20          );
NOTE: Data set "WORKX.want" has 1 observation(s) and 9 variable(s)
21
22          disconnect from odbc;

2

NOTE: Successfully disconnected from database ODBC.
23        quit;
NOTE: Procedure sql step took :
      real time : 0.555
      cpu time  : 0.609


24
25        proc print data=workx.want;
26        run;quit;
NOTE: 1 observations were read from "WORKX.want"
NOTE: Procedure print step took :
      real time : 0.016
      cpu time  : 0.000


27
ERROR: Error printed on page 1

NOTE: Submitted statements took :
      real time : 0.640
      cpu time  : 0.671


/*  _                         _               _   _        _ _      _                            _
| || |     ___ _ __ ___  __ _| |_ ___    __ _| |_| |_ _ __(_) |__  (_)_ __ ___  _ __   ___  _ __| |_
| || |_   / __| `__/ _ \/ _` | __/ _ \  / _` | __| __| `__| | `_ \ | | `_ ` _ \| `_ \ / _ \| `__| __|
|__   _| | (__| | |  __/ (_| | ||  __/ | (_| | |_| |_| |  | | |_) || | | | | | | |_) | (_) | |  | |_
   |_|    \___|_|  \___|\__,_|\__\___|  \__,_|\__|\__|_|  |_|_.__/ |_|_| |_| |_| .__/ \___/|_|   \__|
*/

data workx.fmt;

 length attrib $200;

 set workx.want;

 array nums num_: ;   /*--- num_nam to num_wgt if <7 use character format  ---*/
 array lens len_: ;   /*--- len-nam to len_wgt variable logest lengths     ---*/
 array fmts $8 fmt1-fmt4;

 do over nums;
   colname=scan(vname(nums),2,'_') ;
   putlog colname=;
   select;
     when(nums < 7)  fmts=catx(' ',colname, cats('$',lens,'.')) ; /*--- assume char format $[length] ---*/
     otherwise       fmts=catx(' ',colname, cats(lens,'.')) ;     /*--- assume char format [length] ---*/
   end;
   attrib= catx(' ',attrib,fmts);  /*--- append to create one string            ---*/
   call symputx('attrib',attrib);  /*--- ATTRIB=NAM $7. SEX $1. AGE $2. WGT 3.  ---*/
 end;
 keep attrib fmt:;
run;quit;

proc print data=workx.fmt;
run;quit;

/************************************************************************************/
/* Altair SLC                                                                       */
/*                                                                                  */
/* WORKX.FMT total obs=1                                                            */
/*                                                                                  */
/* Obs                ATTRIB                 FMT1       FMT2       FMT3       FMT4  */
/*                                                                                  */
/*  1     NAM $7. SEX $1. AGE $2. WGT 3.    NAM $7.    SEX $1.    AGE $2.    WGT 3. */
/************************************************************************************/

/*
| | ___   __ _
| |/ _ \ / _` |
| | (_) | (_| |
|_|\___/ \__, |
         |___/
*/

1                                          Altair SLC      12:29 Tuesday, February 10, 2026

NOTE: Copyright 2002-2025 World Programming, an Altair Company
NOTE: Altair SLC 2026 (05.26.01.00.000758)
      Licensed to Roger DeAngelis
NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
NOTE: AUTOEXEC source line
1       +  ï»¿ods _all_ close;
           ^
ERROR: Expected a statement keyword : found "?"
NOTE: Library workx assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\wpswrkx

NOTE: Library slchelp assigned as follows:
      Engine:        WPD
      Physical Name: C:\Progra~1\Altair\SLC\2026\sashelp

NOTE: Library worksas assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\worksas

NOTE: Library workwpd assigned as follows:
      Engine:        WPD
      Physical Name: d:\workwpd


LOG:  12:29:12
NOTE: 1 record was written to file PRINT

NOTE: The data step took :
      real time : 0.024
      cpu time  : 0.000


NOTE: AUTOEXEC processing completed

1         data workx.fmt;
2
3          length attrib $200;
4
5          set workx.want;
6
7          array nums num_: ;   /*--- num_nam to num_wgt if <7 use character format  ---*/
8          array lens len_: ;   /*--- len-nam to len_wgt variable logest lengths     ---*/
9          array fmts $8 fmt1-fmt4;
10
11         do over nums;
12           colname=scan(vname(nums),2,'_') ;
13           putlog colname=;
14           select;
15             when(nums < 7)  fmts=catx(' ',colname, cats('$',lens,'.')) ; /*--- assume char format $[length] ---*/
16             otherwise       fmts=catx(' ',colname, cats(lens,'.')) ;     /*--- assume char format [length] ---*/
17           end;
18           attrib= catx(' ',attrib,fmts);  /*--- append to create one string            ---*/
19           call symputx('attrib',attrib);  /*--- ATTRIB=NAM $7. SEX $1. AGE $2. WGT 3.  ---*/
20         end;
21         keep attrib fmt:;
22        run;

COLNAME=NAM
COLNAME=SEX
COLNAME=AGE

2

COLNAME=WGT
NOTE: 1 observations were read from "WORKX.want"
NOTE: Data set "WORKX.fmt" has 1 observation(s) and 5 variable(s)
NOTE: The data step took :
      real time : 0.053
      cpu time  : 0.015


22      !     quit;
23
24        proc print data=workx.fmt;
25        run;quit;
NOTE: 1 observations were read from "WORKX.fmt"
NOTE: Procedure print step took :
      real time : 0.016
      cpu time  : 0.000


26
ERROR: Error printed on page 1

NOTE: Submitted statements took :
      real time : 0.142
      cpu time  : 0.046

/*___                           _   _        _ _                      _          _        _     _
| ___|   _   _ ___  ___    __ _| |_| |_ _ __(_) |__   _ __ ___   __ _| | _____  | |_ __ _| |__ | | ___
|___ \  | | | / __|/ _ \  / _` | __| __| `__| | `_ \ | `_ ` _ \ / _` | |/ / _ \ | __/ _` | `_ \| |/ _ \
 ___) | | |_| \__ \  __/ | (_| | |_| |_| |  | | |_) || | | | | | (_| |   <  __/ | || (_| | |_) | |  __/
|____/   \__,_|___/\___|  \__,_|\__|\__|_|  |_|_.__/ |_| |_| |_|\__,_|_|\_\___|  \__\__,_|_.__/|_|\___|
*/

data _null_;;

 set workx.fmt;

 put sttrib=;
 call symputx('attrib',attrib);

 rc=dosubl('
    libname xls excel "d:/xls/have.xlsx";

    data workx.class;

       informat &attrib; /*--- NAM $7. SEX $1. AGE $2. WGT 3. ---*/
       format &attrib;   /*--- NAM $7. SEX $1. AGE $2. WGT 3. ---*/

       set xls.class;

    run;

   ');
run;quit;

proc print data=workx.class;
run;quit;

proc contents data=workx.class position;
run;quit;

/**********************************************************************************************************/
/* The CONTENTS Procedure                                             |   WORKX.CLASS total obs=7         */
/*                                                                    |                                   */
/*       List of Variables and Attributes in Creation Order           |   NAM        SEX    AGE    WGT    */
/*                                                                    |                                   */
/*  Number    Variable    Type   Len   Pos    Format   Informat       |   Alfred      0     UU     100    */
/* ________________________________________________________________   |   Alice       1     13     122    */
/*       1    NAM         Char     7     8    $7.      $7.            |   Barbara     F     13     111    */
/*       2    SEX         Char     1    15    $1.      $1.            |   Carol       F     14     113    */
/*       3    AGE         Char     2    16    $2.      $2.            |   Henry       M     14     121    */
/*       4    WGT         Num      8     0    3.        3.            |   James       M     12     116    */
/*                                                                    |   Jane        F     12     109    */
/*                                                                    |                                   */
/* Data Set Name           CLASS                                      |                                   */
/* Member Type             DATA                                       |                                   */
/* Engine                  SAS7BDAT                                   |                                   */
/* Created                 10FEB2026:12:41:30                         |                                   */
/* Last Modified           10FEB2026:12:41:30                         |                                   */
/* Observations                     7                                 |                                   */
/* Variables               4                                          |                                   */
/* Indexes                 0                                          |                                   */
/* Observation Length      24                                         |                                   */
/* Deleted Observations             0                                 |                                   */
/* Data Set Type                                                      |                                   */
/* Label                                                              |                                   */
/* Compressed              NO                                         |                                   */
/* Sorted                  NO                                         |                                   */
/* Data Representation     WINDOWS_64                                 |                                   */
/* Encoding                wlatin1 Windows-1252 Western               |                                   */
/*                                                                    |                                   */
/*           Engine/Host Dependent Information                        |                                   */
/*                                                                    |                                   */
/* Data Set Page Size          4096                                   |                                   */
/* Number of Data Set Pages    1                                      |                                   */
/* First Data Page             1                                      |                                   */
/* Max Obs Per Page            168                                    |                                   */
/* Obs In First Data Page      7                                      |                                   */
/* File Name                   d:\wpswrkx\class.sas7bdat              |                                   */
/* Release Created             9.0101M3                               |                                   */
/* Host Created                XP_PRO                                 |                                   */
/**********************************************************************************************************/

/*
| | ___   __ _
| |/ _ \ / _` |
| | (_) | (_| |
|_|\___/ \__, |
         |___/
*/

1                                          Altair SLC      12:41 Tuesday, February 10, 2026

NOTE: Copyright 2002-2025 World Programming, an Altair Company
NOTE: Altair SLC 2026 (05.26.01.00.000758)
      Licensed to Roger DeAngelis
NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
NOTE: AUTOEXEC source line
1       +  ï»¿ods _all_ close;
           ^
ERROR: Expected a statement keyword : found "?"
NOTE: Library workx assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\wpswrkx

NOTE: Library slchelp assigned as follows:
      Engine:        WPD
      Physical Name: C:\Progra~1\Altair\SLC\2026\sashelp

NOTE: Library worksas assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\worksas

NOTE: Library workwpd assigned as follows:
      Engine:        WPD
      Physical Name: d:\workwpd


LOG:  12:41:29
NOTE: 1 record was written to file PRINT

NOTE: The data step took :
      real time : 0.017
      cpu time  : 0.000


NOTE: AUTOEXEC processing completed

1         data _null_;;
2
3          set workx.fmt;
4
5          put sttrib=;
6          call symputx('attrib',attrib);
7
8          rc=dosubl('
9             libname xls excel "d:/xls/have.xlsx";
10
11            data workx.class;
12               informat &attrib; /*--- NAM $7. SEX $1. AGE $2. WGT 3. ---*/
13               format &attrib;   /*--- NAM $7. SEX $1. AGE $2. WGT 3. ---*/
14               set xls.class;
15            run;
16
17           ');
18        run;quit;
NOTE: Variable "STTRIB" may not be initialized

STTRIB=.
19            libname xls excel "d:/xls/have.xlsx";    data workx.class;
informat &attrib; /*--- NAM $7. SEX $1. AGE $2. WGT 3. ---*/
format &attrib;   /*--- NAM $7. SEX $1. AGE $2. WGT 3. ---*/
set xls.class;    run;
NOTE: Library xls assigned as follows:
      Engine:        OLEDB
      Physical Name: d:/xls/have.xlsx


2


NOTE: 7 observations were read from "XLS.class"
NOTE: Data set "WORKX.class" has 7 observation(s) and 4 variable(s)
NOTE: The data step took :
      real time : 0.063
      cpu time  : 0.062


ERROR: Error printed on page 1

NOTE: Submitted statements took :
      real time : 0.504
      cpu time  : 0.671
NOTE: 1 observations were read from "WORKX.fmt"
NOTE: The data step took :
      real time : 0.504
      cpu time  : 0.671


20
21        proc print data=workx.class;
22        run;quit;
NOTE: 7 observations were read from "WORKX.class"
NOTE: Procedure print step took :
      real time : 0.004
      cpu time  : 0.015


23
24        proc contents data=workx.class position;
25        run;quit;
NOTE: Procedure contents step took :
      real time : 0.079
      cpu time  : 0.046



ERROR: Error printed on page 1

NOTE: Submitted statements took :
      real time : 0.774
      cpu time  : 0.781

/*__                                                                _
 / /_    __ _ _   _  ___ _ __ _   _  _ __   __ _ _ __ ___   ___  __| |   _ __ __ _ _ __   __ _  ___
| `_ \  / _` | | | |/ _ \ `__| | | || `_ \ / _` | `_ ` _ \ / _ \/ _` |__| `__/ _` | `_ \ / _` |/ _ \
| (_) || (_| | |_| |  __/ |  | |_| || | | | (_| | | | | | |  __/ (_| |__| | | (_| | | | | (_| |  __/
 \___/  \__, |\__,_|\___|_|   \__, ||_| |_|\__,_|_| |_| |_|\___|\__,_|  |_|  \__,_|_| |_|\__, |\___|
           |_|                |___/                                                      |___/
*/

/*--- note simpler class instead of `class$`. Class acts as a table (named-range rather than a sheet) ---*/

proc sql;
  connect to odbc (dsn="class");

  create
     table workx.want as
  select
     *
  from connection to odbc (
    select
      count(*)                       as rows
     ,sum(iif(isnumeric(NAM),1, 0))  as num_nam
     ,sum(iif(isnumeric(SEX),1, 0))  as num_sex
     ,sum(iif(isnumeric(AGE),1, 0))  as num_age
     ,sum(iif(isnumeric(WGT),1, 0))  as num_wgt
     ,max(len(nam))                  as len_nam
     ,max(len(sex))                  as len_sex
     ,max(len(age))                  as len_age
     ,max(len(wgt))                  as len_wgt
    from class
  );

  disconnect from odbc;
quit;

proc print data=workx.want;
run;quit;


/************************************************************************************/
/* Altair SLC                                                                       */
/*                                                                                  */
/* WORKX.WANT t otal obs=1                                                            */
/*                                                                                  */
/* Obs                ATTRIB                 FMT1       FMT2       FMT3       FMT4  */
/*                                                                                  */
/*  1     NAM $7. SEX $1. AGE $2. WGT 3.    NAM $7.    SEX $1.    AGE $2.    WGT 3. */
/************************************************************************************/

/*
| | ___   __ _
| |/ _ \ / _` |
| | (_) | (_| |
|_|\___/ \__, |
         |___/
*/

1                                          Altair SLC      12:55 Tuesday, February 10, 2026

NOTE: Copyright 2002-2025 World Programming, an Altair Company
NOTE: Altair SLC 2026 (05.26.01.00.000758)
      Licensed to Roger DeAngelis
NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
NOTE: AUTOEXEC source line
1       +  ï»¿ods _all_ close;
           ^
ERROR: Expected a statement keyword : found "?"
NOTE: Library workx assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\wpswrkx

NOTE: Library slchelp assigned as follows:
      Engine:        WPD
      Physical Name: C:\Progra~1\Altair\SLC\2026\sashelp

NOTE: Library worksas assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\worksas

NOTE: Library workwpd assigned as follows:
      Engine:        WPD
      Physical Name: d:\workwpd


LOG:  12:55:10
NOTE: 1 record was written to file PRINT

NOTE: The data step took :
      real time : 0.018
      cpu time  : 0.031


NOTE: AUTOEXEC processing completed

1         proc sql;
2           connect to odbc (dsn="class");
NOTE: Connected to DB: class (EXCEL version 12.00.0000)
NOTE: Connected to DB: class (EXCEL version 12.00.0000)
NOTE: Successfully connected to database ODBC as alias ODBC.
3
4           create
5              table workx.want as
6           select
7              *
8           from connection to odbc (
9             select
10              count(*)                       as rows
11             ,sum(iif(isnumeric(NAM),1, 0))  as num_nam
12             ,sum(iif(isnumeric(SEX),1, 0))  as num_sex
13             ,sum(iif(isnumeric(AGE),1, 0))  as num_age
14             ,sum(iif(isnumeric(WGT),1, 0))  as num_wgt
15             ,max(len(nam))                  as len_nam
16             ,max(len(sex))                  as len_sex
17             ,max(len(age))                  as len_age
18             ,max(len(wgt))                  as len_wgt
19            from class
20          );
NOTE: Data set "WORKX.want" has 1 observation(s) and 9 variable(s)
21
22          disconnect from odbc;

2

NOTE: Successfully disconnected from database ODBC.
23        quit;
NOTE: Procedure sql step took :
      real time : 0.634
      cpu time  : 0.640


24
25        proc print data=workx.want;
26        run;quit;
NOTE: 1 observations were read from "WORKX.want"
NOTE: Procedure print step took :
      real time : 0.008
      cpu time  : 0.015


27
ERROR: Error printed on page 1

NOTE: Submitted statements took :
      real time : 0.717
      cpu time  : 0.734


/*____                                           _
|___  |   __ _ _   _  ___ _ __   _   _ _   _ ___(_)_ __   __ _  _ __  _ __ ___   ___   _ __
   / /   / _` | | | |/ _ \ `__| | | | | | | / __| | `_ \ / _` || `_ \| `__/ _ \ / __| | `__|
  / /   | (_| | |_| |  __/ |    | |_| | |_| \__ \ | | | | (_| || |_) | | | (_) | (__  | |
 /_/     \__, |\__,_|\___|_|     \__, |\__,_|___/_|_| |_|\__, || .__/|_|  \___/ \___| |_|
            |_|                  |___/                   |___/ |_|
*/
/*--- using win 11 odbc with r package RODBC ---*/

options set=RHOME "C:\Progra~1\R\R-4.5.2\bin\r";
proc r;
submit;
library(RODBC);
ch <- odbcConnect("class");
want <- sqlQuery(ch,"
    select
      count(*)                         as rows
     ,sum(iif(isnumeric([NAM]),1, 0))  as num_nam
     ,sum(iif(isnumeric([SEX]), 1, 0)) as num_sex
     ,sum(iif(isnumeric([AGE]), 1, 0)) as num_age
     ,sum(iif(isnumeric([WGT]), 1, 0)) as num_wgt
     ,max(len(nam))                    as len_nam
     ,max(len(sex))                    as len_sex
     ,max(len(age))                    as len_age
     ,max(len(wgt))                    as len_wgt
    from
      class
    ");
want
endsubmit;
import data=workx.final r=want;
run;

proc print data=workx.final;
run;


/************************************************************************************/
/* Altair SLC                                                                       */
/*                                                                                  */
/* WORKX.FINAL otal obs=1                                                           */
/*                                                                                  */
/* Obs                ATTRIB                 FMT1       FMT2       FMT3       FMT4  */
/*                                                                                  */
/*  1     NAM $7. SEX $1. AGE $2. WGT 3.    NAM $7.    SEX $1.    AGE $2.    WGT 3. */
/************************************************************************************/
/*___        _                       _                       _                                          _          _ _
 ( _ )    __| |_ __ ___  _ __     __| | _____      ___ __   | |_ ___   _ __   _____      _____ _ __ ___| |__   ___| | |
 / _ \   / _` | `__/ _ \| `_ \   / _` |/ _ \ \ /\ / / `_ \  | __/ _ \ | `_ \ / _ \ \ /\ / / _ \ `__/ __| `_ \ / _ \ | |
| (_) | | (_| | | | (_) | |_) | | (_| | (_) \ V  V /| | | | | || (_) || |_) | (_) \ V  V /  __/ |  \__ \ | | |  __/ | |
 \___/   \__,_|_|  \___/| .__/   \__,_|\___/ \_/\_/ |_| |_|  \__\___/ | .__/ \___/ \_/\_/ \___|_|  |___/_| |_|\___|_|_|
                        |_|                                           |_|
*/

%macro utl_submit_ps64(
      pgm
     ,return=  /* name for the macro variable from Powershell */
     )/des="Semi colon separated set of python commands - drop down to python";


  /*
      %let pgm='Get-Content -Path d:/txt/back.txt | Measure-Object -Line | clip;';
  */

  * write the program to a temporary file;
  filename py_pgm "%sysfunc(pathname(work))/py_pgm.ps1" lrecl=32766 recfm=v;
  data _null_;
    length pgm  $32755 cmd $1024;
    file py_pgm ;
    pgm=&pgm;
    semi=countc(pgm,';');
      do idx=1 to semi;
        cmd=cats(scan(pgm,idx,';'));
        if cmd=:'. ' then
           cmd=trim(substr(cmd,2));
         put cmd $char384.;
         putlog cmd $char384.;
      end;
  run;quit;
  %let _loc=%sysfunc(pathname(py_pgm));
  %put &_loc;
  filename rut pipe  "powershell.exe -executionpolicy bypass -file &_loc ";
  data _null_;
    file print;
    infile rut;
    input;
    put _infile_;
    putlog _infile_;
  run;
  filename rut clear;
  filename py_pgm clear;

  * use the clipboard to create macro variable;
  %if "&return" ^= "" %then %do;
    filename clp clipbrd ;
    data _null_;
     length txt $200;
     infile clp;
     input;
     putlog "*******  " _infile_;
     call symputx("&return",_infile_,"G");
    run;quit;
  %end;

%mend utl_submit_ps64;

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
