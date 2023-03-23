deac wind all
set cent on
set mouse off
CLOSE ALL
set status bar off
SET SAFETY OFF
SET ESCAPE OFF
SET DEVI TO SCRE
SET COLO TO W+/B
set sysmenu off
set talk off
SET DATE TO BRITISH
set clock to 0,69

clea
onam=space(20)
unam=space(10)
pwrd=space(6)
allok="n"
machine=1

use mach
go top
if upper(mach)="SUZUKI"
store 1 to machine
else
if upper(mach)="BASIC"
store 2 to machine
else
if upper(mach)="ALPINE"
store 3 to machine
else
store 1 to machine
ENDIF
endif
endif
use

a=0
do while a<7
clea
@a,2*a to 24-a,79-2*a panel 
a=a+1
a1=0
do while a1<6000
a1=a1+1
enddo
enddo

define wind wino1 from 0,0 to 23,79 
define wind wino from 0,0 to 24,79 footer "Press <Ctrl>W to Select and Print the Weighslip"  
DEFINE WIND WIN1 FROM 0,0 TO 24,79 
define wind win113 from 0,0 to 24,79
define wind win111 from 0,60 to 10,79
define wind win112 from 0,60 to 2,79
acti wind win113

SET COLO TO W+/R

@ 7,14 TO 17,65 CLEA
@ 11,17 say "  Weighing program designed and installed by:"
@ 13,22 say "    DADHWAL WEIGHING SYSTEMS"
SET COLO TO W+/B

****
@1,25 say "  Enter Username :        "
@1,44 get unam
read
@2,25 say "  Enter Password :        "
ctr=0
@2,44 say ""
do while ctr<6
READ
@ 2,44+ctr say "*"
pchar=lastkey()
pwrd=pwrd+chr(pchar)
ctr=ctr+1
if lastkey()=13
ctr=6
endif
enddo
pwrd=ltrim(pwrd)

use oper
go top
locate for uname=unam and pword=pwrd
if found()
store oname to onam
allok="y" 
use logs
go top
locate for oname=onam and dat=date()
if not found()
append blank
replace dat with date()
replace oname with onam
replace login with left(time(),5)
replace logout with left(time(),5)
endif
endif
close all

DEFINE MENU MAIN 
DEFINE PAD A OF MAIN PROMPT "SIMPLE" AT 8,2 
DEFINE PAD B OF MAIN PROMPT "SPECIAL" AT 8,12
DEFINE PAD F OF MAIN PROMPT "TRANSPORT" AT 8,23
DEFINE PAD C OF MAIN PROMPT "CUSTOMER DETAILS" AT 8,35
DEFINE PAD D OF MAIN PROMPT "DAILY REPORT" AT 8,55 
DEFINE PAD E OF MAIN PROMPT "QUIT" AT 8,71 
ON PAD A OF MAIN acti popup SITI
ON PAD B OF MAIN ACTI POPUP Coal
ON PAD F OF MAIN ACTI POPUP MATR		
ON PAD C OF MAIN ACTI POPUP CUDE
ON PAD D OF MAIN ACTI POPUP DARE
ON PAD E OF MAIN ACTI POPUP QUT

DEFINE POPUP MATR FROM 10,23 shadow 
DEFINE BAR 1 OF MATR PROMPT "FIRST WEIGHT"
DEFINE BAR 2 OF MATR PROMPT "SECOND WEIGHT"
DEFINE BAR 3 OF MATR PROMPT "VIEW RECORDS"
DEFINE BAR 4 OF MATR PROMPT "TRANSPORT REPORTS"
DEFINE BAR 5 OF MATR PROMPT "MATERIALS"
ON SELE BAR 1 OF matr acti popup mafwpo
oN SELE BAR 2 OF matr DO MASISW
ON SELE BAR 3 OF matr DO MASIVF
ON BAR 4 OF matr Acti POPUP MAREPO
ON BAR 5 OF matr ACTI POPUP MAMATE

DEFINE POPUP maFWPO FROM 10,40 
DEFINE BAR 1 OF maFWPO PROMPT "NEW FIRST WEIGHT "
define bar 2 of maFWPO prompt "RECALL OLD WEIGHT"
ON SELE BAR 1 OF maFWPO do MAsifw
ON SELE BAR 2 OF maFWPO DO maSIRFW  
 
DEFINE POPUP MAMATE FROM 15,40 
DEFINE BAR 1 OF MAMATE PROMPT "ADDING MATERIAL"
define bar 2 of MAMATE prompt "MODIFY MATERIAL"
define bar 3 of MAMATE prompt "DELETE MATERIAL"
ON SELE BAR 1 OF MAMATE DO MAAD
ON SELE BAR 2 OF MAMATE DO MAMO  
ON SELE BAR 3 OF MAMATE DO MADE  

DEFINE POPUP MAREPO FROM 13,40
DEFINE BAR 1 OF MAREPO PROMPT "DATE  -  WISE"
define bar 2 of MArepo prompt "OLD WEIGHSLIP"
define bar 3 of MArepo prompt "SHIFTS - WISE"
ON SELE BAR 1 OF MAREPO DO MAREDW
ON SELE BAR 2 OF MAREPO DO MAREOW  
ON SELE BAR 3 OF MAREPO DO MARESW  


DEFINE POPUP coal FROM 10,12 shadow
DEFINE BAR 1 OF COAL PROMPT "FIRST WEIGHT"
DEFINE BAR 2 OF COAL PROMPT "SECOND WEIGHT"
DEFINE BAR 3 OF COAL PROMPT "VIEW FILE"
ON SELE BAR 1 OF COAL DO COFW 
ON SELE BAR 2 OF COAL DO COSW 
ON SELE BAR 3 OF COAL DO COVF 

DEFINE POPUP SITI FROM 10,2 shadow
DEFINE BAR 1 OF SITI PROMPT "FIRST WEIGHT  "
DEFINE BAR 2 OF SITI PROMPT "SECOND WEIGHT "
DEFINE BAR 3 OF SITI PROMPT "VIEW FILE     "
DEFINE BAR 4 OF SITI PROMPT "DAILY REPORT  "
ON SELE BAR 1 OF SITI DO SIFW 
ON SELE BAR 2 OF SITI DO SISW 
ON SELE BAR 3 OF SITI DO SIVF 
ON BAR 4 OF SITI acti popup sdare 

DEFINE POPUP CUDE FROM 10,35 shadow 
DEFINE BAR 1 OF CUDE PROMPT "NEW ENTRY   "
DEFINE BAR 2 OF CUDE PROMPT "ADD QUANTITY"
DEFINE BAR 3 OF CUDE PROMPT "VIEW RECORDS"
DEFINE BAR 4 OF CUDE PROMPT "PRINT       "
DEFINE BAR 5 OF CUDE PROMPT "COLL./ MAT. "
DEFINE BAR 6 OF CUDE PROMPT "SET PASSWORD"
ON SELE BAR 1 OF CUDE DO CNEW 
ON SELE BAR 2 OF CUDE DO CADQ 
ON SELE BAR 3 OF CUDE DO CMOD 
ON BAR 4 OF CUDE ACTI POPUP CUDEP 
ON BAR 5 OF CUDE ACTI POPUP CAVA
ON SELE BAR 6 OF CUDE DO CSPW 

DEFINE POPUP CUDEP FROM 12,50 shadow 
DEFINE BAR 1 OF CUDEP PROMPT "PRINT ALL RECORDS"
DEFINE BAR 2 OF CUDEP PROMPT "INDIVIDUAL REPORT"
ON SELE BAR 1 OF CUDEP DO CDP 
ON SELE BAR 2 OF CUDEP DO CIND 

define popup dare from 10,55 shadow
DEFINE BAR 1 OF DARE PROMPT "DATE WISE"
DEFINE BAR 2 OF DARE PROMPT "OLD WEIGHSLIP"
DEFINE BAR 3 OF DARE PROMPT "SORTED MATERIAL"
DEFINE BAR 4 OF DARE PROMPT "SORTED PARTY"
DEFINE BAR 5 OF DARE PROMPT "SELECTED MATERIAL"
DEFINE BAR 6 OF DARE PROMPT "SELECTED PARTY"
DEFINE BAR 7 OF DARE PROMPT "OPERATOR SUMMARY"
ON SELE BAR 1 OF DARE DO REDW
ON SELE BAR 2 OF DARE DO REOW
ON SELE BAR 3 OF DARE DO REMW
ON SELE BAR 4 OF DARE DO REPW
ON SELE BAR 5 OF DARE DO RESM
ON SELE BAR 6 OF DARE DO RESP
ON SELE BAR 7 OF DARE DO REOS

define popup sdare from 14,15 shadow
DEFINE BAR 1 OF SDARE PROMPT "DATE WISE"
DEFINE BAR 2 OF SDARE PROMPT "OLD WEIGHSLIP"
ON SELE BAR 1 OF SDARE DO SREDW
ON SELE BAR 2 OF SDARE DO SREOW

DEFINE POPUP CAVA FROM 13,50 shadow 
DEFINE BAR 1 OF CAVA PROMPT "NEW ENTRY"
DEFINE BAR 2 OF CAVA PROMPT "VIEW RECORDS"
DEFINE BAR 3 OF CAVA PROMPT "PRINT"
ON SELE BAR 1 OF CAVA DO CAD 
ON SELE BAR 2 OF CAVA DO CMO 
ON SELE BAR 3 OF CAVA DO CDE 

DEFINE POPUP QUT FROM 10,71 shadow 
DEFINE BAR 1 OF QUT PROMPT "QUIT"
ON SELE BAR 1 OF QUT DO QUIT 

on key label f1 do ee
on key label f2 do ee
on key label f3 do ee
on key label f4 do ee
on key label f5 do ee
on key label f6 do ee
on key label f7 do ee
on key label f8 do ee
on key label f9 do ee
on key label f10 do ee
on key label f11 do ee
on key label f12 do ee
on key label esc do eee
on key label shift+f2 do party
on key label shift+f4 do testing
on key label shift+f3 do creq
on key label shift+f5 do oppw

if allok="y" and onam="manager"
	@3,25 say "  Operator       : "+onam
	@21,25 say "Press any key to continue"
	wait""
	do admin
else 
	if allok="y"
		@3,25 say "  Operator       : "+onam
		@21,25 say "Press any key to continue"
		wait ""
		clea
		DO HEAD
		do tail
		@22,0 SAY "    Make sure the platform is clear before every Tare and Gross weighment.    " colo w+/r*
		@6,0 TO 21,78
		ACTI MENU MAIN
		DEAC WIND ALL
		DEAC MENU MAIN
		CLEA
	else
		@20,8 say "You may be an unauthorized user...check your password and retry"
		wait ""
		deac wind all
		clea
	endif
endif

return 

proc eee
close all
rele all
return

PROC EE
RETURN

PROC MASIVF
ACTI WIND WIN1
USE SIMPLET
go bottom
BROW noappend noedit nodelete
USE
DEAC WIND win1
RETURN

PROC MAAD
USE MATERIAL
APPEND
USE
RETURN

PROC MAMO
USE MATERIAL
BROW
USE
RETURN

PROC MADE
ON KEY LABEL ESC DO EE
ACTI WIND WINOO
ACTI WIND WINO1
USE MATERIAL
BROW
CH="n"
@12,12 SAY "ARE YOU SURE TO DELETE THE SELECTED RECORD(Y/N)? : "GET CH
READ
IF UPPER(CH)="Y"
DELETE
PACK
ENDIF
CLOSE ALL
DEAC WIND WINO1
DEAC WIND WINOO
ON KEY LABEL ESC DO EEE
RETURN

PROC MASIFW
ACTI WIND WIN1
FW=0
***** READING COM PORT********************
acti wind win111
close all
@0,0 say machine
if machine=1 
run wtread
b=fopen("weight.dat",10)
endif
if machine=2 
!basica ccc
b=fopen("cc.dbf",10)
endif
if machine=3 
b=fopen("cc.dbf",10)
endif
a=fread(b,10)
clea
store val(a) to fW
CLOSE ALL
deac wind win111
*****************************************
if fw<21

CLEA
DO HEAD
SEA=SPACE(9)
SE=0
DASE=0
CS=SPACE(25)
CN=SPACE(25)
AD=SPACE(25)
MN=space(15)
VN=SPACE(15)
@ 6,16 SAY "YEAR         :"
@ 7,16 say "SHIFT        :"
@ 8,16 SAY "DAILY SERIAL :"
@ 9,16 SAY "DATE         :"
@10,16 SAY "TIME         :"
@12,16 SAY "VEHICLE NO.  :"
@13,16 say "NAME         :"
@14,16 SAY "ADD/ DEST.   :"
@15,16 say "CHALLAN/SO.NO:"
@16,16 SAY "MATERIAL     :"

IF MONTH(DATE())>3
STORE RIGHT(STR(YEAR(DATE())),4)+"-"+RIGHT(STR(YEAR(DATE())+1),4) TO SEA
ELSE
STORE RIGHT(STR(YEAR(DATE())-1),4)+"-"+RIGHT(STR(YEAR(DATE())),4) TO SEA
ENDIF

shi=space(1)
if val(time())>=8 and val(time())<16
shi="A"
else
if val(time())>=16 and val(time())<24
shi="B"
else
if val(time())>=0 and val(time())<=8
shi="C"
endif
endif
endif

USE SIMPLET
GO BOTTOM
IF DATE_IN=DATE()
STORE (D_SERIAL+1) TO DASE
ELSE
STORE 1 TO DASE
ENDIF
close all
@ 6,31 SAY SEA
@ 7,31 SAY SHI
@ 8,31 SAY DASE PICT '99999'
@ 9,31 SAY DATE()
@10,31 SAY LEFT(TIME(),5)
@12,31 GET VN
@13,31 GET CN
@14,31 GET AD
@15,31 GET CS
read
use material
define popup abc prompt field material from 16,31
acti popup abc
store material to mn
use
@16,31 say mn

@17,16 SAY "FIRST WEIGHT : Press any key twice to weigh"
wait ""
**** READING COM PORT********************
acti wind win111
close all
if machine=1 
run wtread
b=fopen("weight.dat",10)
endif
if machine=2 
!basica ccc
b=fopen("cc.dbf",10)
endif
if machine=3 
b=fopen("cc.dbf",10)
endif
a=fread(b,10)
clea
store val(a) to fW
CLOSE ALL
deac wind win111
*****************************************
@17,31 say "                               "
@17,31 SAY FW PICT '9999999'+" Kg"
OK="y"
@ 21,15 SAY "IS DATA OK (Y/N)? "GET OK
READ
IF OK="y" OR OK="Y"
USE SIMPLET
APPEND BLANK
REPLACE SEASON WITH SEA 
REPLACE D_SERIAL WITH DASE
REPLACE SHIFT WITH SHI
REPLACE DATE_IN WITH DATE()
REPLACE TIME_IN WITH LEFT(TIME(),5)
REPLACE C_NAME WITH CN
REPLACE C_ADDRESS WITH AD
REPLACE CHALLAN_SO WITH CS
REPLACE MATERIAL WITH MN
REPLACE V_NO WITH VN
REPLACE FIRST_WT WITH FW
CLOSE ALL
ppname=space(30)
ppinf=space(30)
ppadd=space(30)
use party
go top
store pname to ppname
store pinf to ppinf
store padd to ppadd
close all
CHO=SPACE(1)
@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
READ
IF CHO="P" OR CHO="p"
do while cho="p" or cho="P"
clea
set printer on
set device to print
set printer to
? chr(18)+chr(14)+space(5)+ppname
? chr(14)+SPACE(5)+ppinf
? chr(14)+SPACE(5)+ppadd
?
? chr(18)+space(27)+"YEAR   : "+sea
? chr(18)+SPACE(27)+"DATE   : "+DTOC(DATE())
? chr(18)+SPACE(27)+"SHIFT  : "+SHI
? SPACE(27)+"TIME IN: "+LEFT(TIME(),5)
?
?
? SPACE(20)+"DAILY SERIAL : "+STR(DASE,6)
?
? SPACE(20)+"VEHICLE NO   : "+VN
? SPACE(20)+"NAME         : "+CN
? SPACE(20)+"ADDRESS      : "+AD
? SPACE(20)+"CHALLAN/SO.NO: "+CS
? SPACE(20)+"MATERIAL NAME: "+MN 
?
? chr(14)+SPACE(5)+"FIRST WEIGHT :"+STR(FW,10)+" Kg" 
?
?
? chr(18)+"                                                    Weighing Operator" 
? "---------------------------------------------------------------------"
?
?
set printer to
set device to screen
set printer off
cho=space(1)
@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
READ
enddo
WAIT " "
ENDIF
ELSE
@ 21,10 SAY "THIS CUSTOMER DOES NOT EXIST"
ENDIF
CLEA
else
@10,20 say "There is load on the platform"
@11,20 say "Make sure the platform is clear..."
wait ""
endif
DEAC WIND WIN1
DO HEAD
do tail
SET COLO TO W+/B
close all
rele all
close all
RETURN


PROC MASIRFW
ACTI WIND WIN1
CS=SPACE(25)
CN=SPACE(25)
AD=SPACE(25)
MN=space(15)
VN=SPACE(15)

CLEA
SIDAT=DATE()
SISER=0
@ 10,10 SAY "ENTER DATE          : "GET SIDAT
READ
@ 11,10 SAY "ENTER SERIAL NUMBER : "GET SISER
READ
CLOSE ALL
USE SIMPLET
GO TOP
LOCATE FOR DATE_IN=SIDAT AND D_SERIAL=SISER
IF FOUND()

STORE V_NO TO VN
STORE FIRST_WT TO FW
STORE C_NAME TO CN
STORE C_ADDRESS TO AD
STORE CHALLAN_SO TO CS
STORE MATERIAL TO MN
 
CLEA
DO HEAD
SEA=SPACE(9)
SE=0
DASE=0
@ 6,16 SAY "YEAR         :"
@ 7,16 say "SHIFT        :"
@ 8,16 SAY "DAILY SERIAL :"
@ 9,16 SAY "DATE         :"
@10,16 SAY "TIME         :"
@12,16 SAY "VEHICLE NO.  :"
@13,16 say "NAME         :"
@14,16 SAY "ADD/ DEST.   :"
@15,16 say "CHALLAN/SO.NO:"
@16,16 SAY "MATERIAL     :"
@17,16 SAY "FIRST WEIGHT :"

IF MONTH(DATE())>3
STORE RIGHT(STR(YEAR(DATE())),4)+"-"+RIGHT(STR(YEAR(DATE())+1),4) TO SEA
ELSE
STORE RIGHT(STR(YEAR(DATE())-1),4)+"-"+RIGHT(STR(YEAR(DATE())),4) TO SEA
ENDIF

shi=space(1)
if val(time())>=6 and val(time())<14
shi="A"
else
if val(time())>=14 .and. val(time())<22
shi="B"
else
if val(time())>=22 .and. val(time())<=23
shi="C"
else
if val(time())>=0 and val(time())<6
shi="C"
endif
endif
endif
endif
CLOSE ALL
USE SIMPLET
GO BOTTOM
IF DATE_IN=DATE()
STORE (D_SERIAL+1) TO DASE
ELSE
STORE 1 TO DASE
ENDIF
close all
@ 6,31 SAY SEA
@ 7,31 SAY SHI
@ 8,31 SAY DASE PICT '99999'
@ 9,31 SAY DATE()
@10,31 SAY LEFT(TIME(),5)
@12,31 SAY VN
@13,31 SAY CN
@14,31 SAY AD
@15,31 SAY CS
@16,31 say mn
@17,31 SAY FW PICT '9999999'+" Kg"
OK="y"
@ 21,15 SAY "IS DATA OK (Y/N)? "GET OK
READ
IF OK="y" OR OK="Y"
USE SIMPLET
APPEND BLANK
REPLACE SEASON WITH SEA 
REPLACE D_SERIAL WITH DASE
REPLACE SHIFT WITH SHI
REPLACE DATE_IN WITH DATE()
REPLACE TIME_IN WITH LEFT(TIME(),5)
REPLACE C_NAME WITH CN
REPLACE C_ADDRESS WITH AD
REPLACE CHALLAN_SO WITH CS
REPLACE MATERIAL WITH MN
REPLACE V_NO WITH VN
REPLACE FIRST_WT WITH FW
CLOSE ALL
ppname=space(30)
ppinf=space(30)
ppadd=space(30)
use party
go top
store pname to ppname
store pinf to ppinf
store padd to ppadd
close all
CHO=SPACE(1)
@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
READ
IF CHO="P" OR CHO="p"
do while cho="p" or cho="P"
clea
set printer on
set device to print
set printer to
? chr(18)+chr(14)+space(5)+ppname
? chr(14)+SPACE(5)+ppinf
? chr(14)+SPACE(5)+ppadd
?
? chr(18)+space(27)+"YEAR   : "+sea
? chr(18)+SPACE(27)+"DATE   : "+DTOC(DATE())
? chr(18)+SPACE(27)+"SHIFT  : "+SHI
? SPACE(27)+"TIME IN: "+LEFT(TIME(),5)
?
?
? SPACE(20)+"DAILY SERIAL : "+STR(DASE,6)
?
? SPACE(20)+"VEHICLE NO   : "+VN
? SPACE(20)+"NAME         : "+CN
? SPACE(20)+"ADDRESS      : "+AD
? SPACE(20)+"CHALLAN/SO.NO: "+CS
? SPACE(20)+"MATERIAL NAME: "+MN 
?
? chr(14)+SPACE(5)+"FIRST WEIGHT :"+STR(FW,10)+" Kg" 
?
?
? chr(18)+"                                                    Weighing Operator" 
? "---------------------------------------------------------------------"
?
?
set printer to
set device to screen
set printer off
cho=space(1)
@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
READ
enddo
WAIT " "
ENDIF
ELSE
@ 21,10 SAY "THIS CUSTOMER DOES NOT EXIST"
ENDIF
CLEA
ELSE
CLOSE ALL
@ 13,10 SAY "WEIGHMENT DOES NOT EXIST"
@ 14,10 SAY "Press any key to continue...."
wait ""
endif
DEAC WIND WIN1
DO HEAD
do tail
SET COLO TO W+/B
close all
rele all
close all
RETURN



PROC MASISW
ACTI WIND WIN1
FW=0
***** READING COM PORT********************
acti wind win111
close all
if machine=1 
run wtread
b=fopen("weight.dat",10)
endif
if machine=2 
!basica ccc
b=fopen("cc.dbf",10)
endif
if machine=3 
b=fopen("cc.dbf",10)
endif

a=fread(b,10)
clea
store val(a) to fW
CLOSE ALL
deac wind win111
*****************************************
if fw<21

SEA=SPACE(9)
DASE=0
dain=ctod("  /  /  ")
daout=ctod("  /  /  ")
tin=space(5)
tout=space(5)
CN=SPACE(25)
vn=space(15)
AD=SPACE(25)
MN=SPACE(15)
fw=0
sw=0
nw=0
STORE DATE() TO DAIN
@ 9,15 say "                DATE IN : "GET DAIN
READ
@ 11,15 SAY "SERIAL NUMBER TO RECALL : "GET daSE PICT '999999'
READ
SELE 1
USE SIMPLET
LOCATE FOR D_SERIAL=DASE AND DATE_IN=DAIN
IF FOUND()
	IF TIME_OUT<>"     "
		CHH="N"
		ookk="n"
		@13,15 SAY "SECOND WEIGHT ALREADY TAKEN..WANT TO GENERATE NEW SERIAL?"GET OOKK
		READ 
		IF OOKK="Y" OR OOKK="y"
			DO SWGET1
		ELSE
			@14,15 SAY "WEIGHMENT TEMINATED...PRESS ANY KEY TO CONTINUE"
			WAIT ""
		ENDIF 
	ELSE
		DO SWGET
	ENDIF
ELSE
	@ 13,15 SAY "SERIAL NUMBER NOT PRESENT"
	WAIT ""
ENDIF
close all
CLEA
else
@10,20 say "There is load on the platform"
@11,20 say "Make sure the platform is clear..."
wait ""
endif
DEAC WIND WIN1
DO HEAD
do tail
SET COLO TO W+/B
close all
rele all
RETURN


PROC SWGET
store season to sea
store d_serial to dase
STORE SHIFT TO SHI
store date_in to dain
store date() to daout
store time_in to tin
store left(time(),5) to tout
STORE C_NAME TO CN
STORE C_ADDRESS TO AD
STORE CHALLAN_SO TO CS
STORE MATERIAL TO MN
STORE V_NO TO VN
STORE FIRST_WT TO FW
clea

DO HEAD
@ 4,18 say "YEAR   : "
@ 4,28 SAY SEA
@ 4,38 say "DATE   : "
@ 4,48 SAY DATE()
@ 5,28 SAY "TIME   : "
@ 5,38 SAY TIME()
@ 6,28 SAY "SHIFT  : "
@ 6,38 SAY SHI
@ 8,22 SAY "DAILY SERIAL :"
@ 9,22 SAY "TIME IN      :"
@ 10,22 SAY "TIME OUT     :"
@ 12,22 SAY "VEHICLE NO.  :"
@ 13,22 SAY "NAME         :"
@ 14,22 SAY "ADDRESS      :"
@ 15,22 SAY "CHALLAN/SO.NO:"
@ 16,22 SAY "MATERIAL NAME:"
@ 17,22 SAY "FIRST WEIGHT :"
@ 8,37 SAY DASE PICT '999999'
@ 9,37 SAY TIN
@10,37 SAY TOUT 
@12,37 SAY VN
@13,37 SAY CN
@14,37 SAY AD
@15,37 SAY CS
@16,37 SAY MN
@17,37 SAY FW PICT '9999999'+" Kg"
@ 18,22 SAY "SECOND WEIGHT: Press any key twice to weigh"
wait ""

**** READING COM PORT********************
acti wind win111
if machine=1 
run wtread
b=fopen("weight.dat",10)
endif
if machine=2 
!basica ccc
b=fopen("cc.dbf",10)
endif
if machine=3 
b=fopen("cc.dbf",10)
endif

a=fread(b,10)
clea
store val(a) to SW
deac wind win111
*****************************************

if fw>sw 
store (fw-sw) to nw
else 
store (sw-fw) to nw
endif

STORE fW TO GROWT
STORE sW TO TARWT

@ 18,22 SAY "SECOND WEIGHT:"
@ 19,22 SAY "NET WEIGHT   :"
@17,37 SAY GROWT PICT '9999999'+" Kg"
@18,37 SAY TARWT PICT '9999999'+" Kg"
@19,37 SAY NW PICT '9999999'+" Kg"
SELE 1
REPLACE DATE_OUT WITH DATE()
REPLACE TIME_OUT WITH LEFT(TIME(),5)
REPLACE SECOND_WT WITH SW 
CLOSE ALL
ppname=space(30)
ppinf=space(30)
ppadd=space(30)
use party
go top
store pname to ppname
store pinf to ppinf
store padd to ppadd
close all
CHO=SPACE(1)
@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
READ
IF CHO="P" OR CHO="p"
do while cho="P" or cho="p"
clea
set printer on
set device to print
set printer to
? chr(18)+chr(14)+space(5)+ppname
? chr(14)+SPACE(5)+ppinf
? chr(14)+SPACE(5)+ppadd
?
? chr(18)+space(27)+"YEAR    : "+sea
? chr(18)+SPACE(27)+"DATE    : "+DTOC(DATE())
? chr(18)+SPACE(27)+"SHIFT   : "+SHI
?
? SPACE(25)+"DAILY SERIAL : "+STR(DASE,6)
? SPACE(25)+"TIME IN      : "+TIN
? SPACE(25)+"TIME OUT     : "+TOUT
? 
? SPACE(25)+"VEHICLE NO   : "+VN
? SPACE(25)+"NAME         : "+CN
? SPACE(25)+"ADDRESS      : "+AD
? SPACE(25)+"CHALLAN/SO.NO: "+CS
? SPACE(25)+"MATERIAL     : "+MN
?
? chr(14)+SPACE(5)+"GROSS WEIGHT : "+STR(GROWT,10)+" Kg" 
? chr(14)+SPACE(5)+"TARE WEIGHT  : "+STR(TARWT,10)+" Kg"
? chr(14)+SPACE(5)+"NET WEIGHT   : "+STR(NW,10)+" Kg"
?
?
? chr(18)+"                                                          Weighing Operator" 
? "-----------------------------------------------------------------------------"
? "N.B:"+chr(15)+"Please check the weight before leaving.                                            Designed & Installed by DADHWAL WEIGHING"
? CHR(18)+"-----------------------------------------------------------------------------"
?
set printer to
set device to screen
set printer off
CHO=SPACE(1)
@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
READ
enddo
WAIT " "
ENDIF
RETURN


PROC SWGET1
store season to sea
store d_serial to dase
STORE SHIFT TO SHI
store date_in to dain
store date() to daout
store time_in to tin
store left(time(),5) to tout
STORE C_NAME TO CN
STORE C_ADDRESS TO AD
STORE CHALLAN_SO TO CS
STORE MATERIAL TO MN
STORE V_NO TO VN
STORE FIRST_WT TO FW
clea

GO BOTTOM
IF DATE_IN=DATE()
STORE (D_SERIAL+1) TO DASE
ELSE
STORE 1 TO DASE
ENDIF

IF MONTH(DATE())>3
STORE RIGHT(STR(YEAR(DATE())),4)+"-"+RIGHT(STR(YEAR(DATE())+1),4) TO SEA
ELSE
STORE RIGHT(STR(YEAR(DATE())-1),4)+"-"+RIGHT(STR(YEAR(DATE())),4) TO SEA
ENDIF
shi=space(1)
if val(time())>=8 and val(time())<16
shi="A"
else
if val(time())>=16 and val(time())<24
shi="B"
else
if val(time())>=0 and val(time())<=8
shi="C"
endif
endif
endif

DO HEAD
@ 4,18 say "YEAR   : "
@ 4,28 SAY SEA
@ 4,38 say "DATE   : "
@ 4,48 SAY DATE()
@ 5,28 SAY "TIME   : "
@ 5,38 SAY TIME()
@ 6,28 SAY "SHIFT  : "
@ 6,38 SAY SHI
@ 8,22 SAY "DAILY SERIAL :"
@ 9,22 SAY "TIME IN      :"
@ 10,22 SAY "TIME OUT     :"
@ 12,22 SAY "VEHICLE NO.  :"
@ 13,22 SAY "NAME         :"
@ 14,22 SAY "ADDRESS      :"
@ 15,22 SAY "CHALLAN/SO.NO:"
@ 16,22 SAY "MATERIAL NAME:"
@ 17,22 SAY "FIRST WEIGHT :"
@ 8,37 SAY DASE PICT '999999'
@ 9,37 SAY TIN
@10,37 SAY TOUT 
@12,37 SAY VN
@13,37 SAY CN
@14,37 SAY AD
@15,37 SAY CS
@16,37 SAY MN
@17,37 SAY FW PICT '9999999'+" Kg"
@ 18,22 SAY "SECOND WEIGHT: Press any key twice to weigh"
wait ""

********* READING COM PORT **************
acti wind win111
if machine=1 
run wtread
b=fopen("weight.dat",10)
endif
if machine=2 
!basica ccc
b=fopen("cc.dbf",10)
endif
if machine=3 
b=fopen("cc.dbf",10)
endif

a=fread(b,10)
clea
store val(a) to SW
deac wind win111
*****************************************

if fw>sw 
store (fw-sw) to nw
else 
store (sw-fw) to nw
endif

STORE fW TO GROWT
STORE sW TO TARWT

@ 17,22 SAY "FIRST WEIGHT :"
@ 17,22 SAY "SECOND WEIGHT:"
@ 19,22 SAY "NET WEIGHT   :"
@17,37 SAY GROWT PICT '9999999'+" Kg"
@18,37 SAY TARWT PICT '9999999'+" Kg"
@19,37 SAY NW PICT '9999999'+" Kg"
NOK="N"
@20,10 say "Is data ok(y/n) : "get nok
read
@20,10 say "Confirm again   : "get nok
read
if nok="y" or nok="Y"
SELE 1
APPEND BLANK
REPLACE SEASON WITH SEA
REPLACE D_SERIAL WITH DASE
REPLACE SHIFT WITH SHI
REPLACE C_NAME WITH CN
REPLACE C_ADDRESS WITH AD
REPLACE CHALLAN_SO WITH CS
REPLACE MATERIAL WITH MN
REPLACE V_NO WITH VN
REPLACE DATE_IN WITH DAIN
REPLACE DATE_OUT WITH DATE()
REPLACE TIME_IN WITH TIN
REPLACE TIME_OUT WITH LEFT(TIME(),5)
REPLACE FIRST_WT WITH FW
REPLACE SECOND_WT WITH SW 
CLOSE ALL
ppname=space(30)
ppinf=space(30)
ppadd=space(30)
use party
go top
store pname to ppname
store pinf to ppinf
store padd to ppadd
close all
CHO=SPACE(1)
@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
READ
IF CHO="P" OR CHO="p"
do while cho="P" or cho="p"
clea
set printer on
set device to print
set printer to
? chr(18)+chr(14)+space(5)+ppname
? chr(14)+SPACE(5)+ppinf
? chr(14)+SPACE(5)+ppadd
?
? chr(18)+space(27)+"YEAR    : "+sea
? chr(18)+SPACE(27)+"DATE    : "+DTOC(DATE())
? chr(18)+SPACE(27)+"SHIFT   : "+SHI
?
? SPACE(25)+"DAILY SERIAL : "+STR(DASE,6)
? SPACE(25)+"TIME IN      : "+TIN
? SPACE(25)+"TIME OUT     : "+TOUT
? 
? SPACE(25)+"VEHICLE NO   : "+VN
? SPACE(25)+"NAME         : "+CN
? SPACE(25)+"ADDRESS      : "+AD
? SPACE(25)+"CHALLAN/SO.NO: "+CS
? SPACE(25)+"MATERIAL     : "+MN
?
? chr(14)+SPACE(5)+"FIRST WEIGHT : "+STR(GROWT,10)+" Kg" 
? chr(14)+SPACE(5)+"SECOND WEIGHT: "+STR(TARWT,10)+" Kg"
? chr(14)+SPACE(5)+"NET WEIGHT   : "+STR(NW,10)+" Kg"
?
?
? chr(18)+"                                                          Weighing Operator" 
? "-----------------------------------------------------------------------------"
? "N.B:"+chr(15)+"Please check the weight before leaving.                                            Designed & Installed by DADHWAL WEIGHING"
? CHR(18)+"-----------------------------------------------------------------------------"
?
set printer to
set device to screen
set printer off
CHO=SPACE(1)
@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
READ
enddo
endif
WAIT " "
ENDIF
RETURN


PROC maREDW
ACTI WIND WIN1
STORE DATE() TO DA1
STORE DATE() TO DA2
@ 12,15 SAY "DATE FROM : "GET DA1
@ 13,15 SAY "DATE TO   : "GET DA2
READ
SELE 1
USE SIMPLET
SET FILTER TO DATE_OUT>=DA1 AND DATE_OUT<=DA2
LN=0
SET PRINTER ON
SET DEVI TO PRINT
SET PRINTER TO
GO TOP
TOT=0
SHI=SPACE(1)
DO WHILE NOT EOF()
SELE 1
STORE SEASON TO SEA
STORE D_SERIAL TO SL
STORE SHIFT TO SHI
STORE DATE_IN TO DIN
STORE DATE_OUT TO DOUT
STORE TIME_IN TO TIN
STORE TIME_OUT TO TOUT
STORE V_NO TO VN
STORE C_NAME TO CN
STORE C_ADDRESS TO AD
STORE CHALLAN_SO TO CS
STORE Material TO MN
STORE FIRST_WT TO FW
STORE SECOND_WT TO SW
IF FW>SW
STORE (FW-SW) TO NW
ELSE 
STORE (SW-FW) TO NW
ENDIF

IF LN=0
WAIT ""
DO HEAD2
ENDIF

TOT=TOT+NW
?CHR(15)+"|"+STR(SL,6)+" |"+TIN+" |"+TOUT+" |"+dtoc(dOUT)+" |"+CN+" |"+vn+" |"+MN+" |"+STR(FW,10)+" |"+STR(SW,10)+" |"+STR(NW,10)+" |"
?CHR(15)+"|"+SPACE(6)+" |"+SPACE(5)+" |"+SPACE(5)+" |"+space(10)+" |"+AD+" |"+space(15)+" |"+SPACE(15)+" |"+SPACE(10)+" |"+SPACE(10)+" |"+SPACE(10)+" |"
SELE 1
LN=LN+1
IF LN>=26
LN=0
ENDIF
SKIP
ENDDO
?CHR(15)+"------------------------------------------------------------------------------------------------------------------------------------"
?CHR(18)+"                                         TOTAL OF NET WEIGHT = "+STR(TOT,10)+" Kg"
?CHR(15)+"------------------------------------------------------------------------------------------------------------------------------------"
? 
?
set printer to
set device to screen
set printer off
SET FILTER TO
WAIT "" 
CLOSE ALL
CLEA
DEAC WIND win1
RETURN



PROC maRESW
ACTI WIND WIN1
SHI=SPACE(1)
STORE DATE() TO DA1
@ 12,15 SAY "DATE    : "GET DA1
@ 13,15 SAY "SHIFT   : "GET SHI
READ
USE SIMPLET
SORT ON MATERIAL,DATE_OUT TO SIMPLET1
USE
SELE 1
USE SIMPLET1
IF UPPER(SHI)="C"
SET FILTER TO (DATE_OUT=DA1 AND VAL(TIME_OUT)>=22) OR (DATE_OUT=DA1+1 AND VAL(TIME_OUT)<6)
else
SET FILTER TO (DATE_OUT=DA1) AND UPPER(SHIFT)=UPPER(SHI)
ENDIF
LN=0
SET PRINTER ON
SET DEVI TO PRINT
SET PRINTER TO

GO TOP
MN1=SPACE(15)
TOT=0
DIN2={}
SHI=SPACE(1)
DO WHILE NOT EOF()
SELE 1
STORE SEASON TO SEA
STORE D_SERIAL TO SL
STORE SHIFT TO SHI
STORE DATE_IN TO DIN
STORE DATE_OUT TO DOUT
STORE TIME_IN TO TIN
STORE TIME_OUT TO TOUT
STORE V_NO TO VN
STORE C_NAME TO CN
STORE C_ADDRESS TO AD
STORE CHALLAN_SO TO CS
STORE Material TO MN
STORE FIRST_WT TO FW
STORE SECOND_WT TO SW
IF FW>SW
STORE (FW-SW) TO NW
ELSE 
STORE (SW-FW) TO NW
ENDIF

IF LN=0
WAIT ""
DO HEAD22
ENDIF

TOT=TOT+NW
?CHR(15)+"|"+STR(SL,6)+" |"+TIN+" |"+TOUT+" |"+dtoc(dOUT)+" |"+CN+" |"+vn+" |"+MN+" |"+STR(FW,10)+" |"+STR(SW,10)+" |"+STR(NW,10)+" |"
?CHR(15)+"|"+SPACE(6)+" |"+SPACE(5)+" |"+SPACE(5)+" |"+space(10)+" |"+AD+" |"+space(15)+" |"+SPACE(15)+" |"+SPACE(10)+" |"+SPACE(10)+" |"+SPACE(10)+" |"
SELE 1
LN=LN+1
IF LN>=26
LN=0
ENDIF
SKIP
SELE 1
STORE MATERIAL TO MN1
IF MN<>MN1
?CHR(15)+"------------------------------------------------------------------------------------------------------------------------------------"
?CHR(18)+"                      TOTAL OF MATERIAL "+MN+" = "+STR(TOT,10)+" Kg"
?CHR(15)+"------------------------------------------------------------------------------------------------------------------------------------"
? 
LN=LN+2
TOT=0
ENDIF

ENDDO
?
?
SET PRINTER TO
SET DEVI TO SCRE
SET PRINTER OFF
SET FILTER TO
WAIT "" 
CLOSE ALL
CLEA
DEAC WIND win1
RETURN




PROC maREOW
on key label esc do ee
ACTI WIND WIN1
USE SIMPLET
acti wind wino
acti wind wino1
GO BOTTOM
BROW FIELDS SEASON,D_SERIAL,DATE_IN,DATE_OUT,TIME_IN,TIME_OUT,V_NO,C_NAME,C_ADDRESS,CHALLAN_SO,Material,FIRST_WT,SECOND_WT NOEDIT NOAPPEND NODELETE

deac wind wino1
deac wind wino
STORE SEASON TO SEA
STORE D_SERIAL TO DASE
STORE SHIFT TO SHI
STORE DATE_IN TO DIN
STORE DATE_OUT TO DOUT
STORE TIME_IN TO TIN
STORE TIME_OUT TO TOUT
STORE V_NO TO VN
STORE C_NAME TO CN
STORE C_ADDRESS TO AD
STORE CHALLAN_SO TO CS
STORE Material TO MN
STORE FIRST_WT TO FW
STORE SECOND_WT TO SW
if fw>sw 
store (fw-sw) to nw
STORE FW TO GROWT
STORE SW TO TARWT
else 
store (sw-fw) to nw
STORE SW TO GROWT
STORE FW TO TARWT
endif

CLOSE ALL

USE PARTY 
GO TOP
store pname to ppname
store pinf to ppinf
store padd to ppadd
close all

CHO=SPACE(1)
@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
READ
IF CHO="P" OR CHO="p"
do while cho="P" or cho="p"
clea
set printer on
set device to print
set printer to
? chr(18)+chr(14)+space(5)+ppname
? chr(14)+SPACE(5)+ppinf
? chr(14)+SPACE(5)+ppadd
?
? chr(18)+space(27)+"YEAR    : "+sea
? chr(18)+SPACE(27)+"DATE    : "+DTOC(DOUT)
? chr(18)+SPACE(27)+"SHIFT   : "+SHI
?
? SPACE(25)+"DAILY SERIAL : "+STR(DASE,6)
? SPACE(25)+"TIME IN      : "+TIN
? SPACE(25)+"TIME OUT     : "+TOUT
? 
? SPACE(25)+"VEHICLE NO   : "+VN
? SPACE(25)+"NAME         : "+CN
? SPACE(25)+"ADDRESS      : "+AD
? SPACE(25)+"CHALLAN/SO.NO: "+CS
? SPACE(25)+"MATERIAL     : "+MN
?
? chr(14)+SPACE(5)+"GROSS WEIGHT : "+STR(GROWT,10)+" Kg" 
? chr(14)+SPACE(5)+"TARE WEIGHT  : "+STR(TARWT,10)+" Kg"
? chr(14)+SPACE(5)+"NET WEIGHT   : "+STR(NW,10)+" Kg"
?
?
? chr(18)+"                                                          Weighing Operator" 
? "-----------------------------------------------------------------------------"
? "N.B:"+chr(15)+"Please check the weight before leaving.                                            Designed & Installed by DADHWAL WEIGHING"
? CHR(18)+"-----------------------------------------------------------------------------"
?
set printer to
set device to screen
set printer off
CHO=SPACE(1)
@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
READ
enddo
ENDIF
WAIT ""
CLEA
DEAC WIND WIN1
CLEA
DO HEAD
@ 6,24 to 20,52 DOUBLE
SET COLO TO W+/W
@ 7,26 TO 19,50 CLEA
SET COLO TO W+/B
on key label esc do eee
RETURN




PROC CAD
ACTI WIND WIN1
CH="y"
DO WHILE CH="y" OR CH="Y"
CLEA
@ 3,25 SAY "..... COLLIERY/ MATERIAL ENTRY SCREEN ....."
MC=SPACE(8)
MN=SPACE(15)
@ 6,20 SAY "COLL./MAT. CODE  "
@ 7,20 SAY "COLY/MAT NAME    "

do while len(alltrim(mc))<=2
MC=SPACE(8)
@6,40 GET mc
read
if len(alltrim(mc))<=2
@20,10 say "Please enter a valid code "
wait ""
@20,10 say "                          "
endif
enddo

IF MC=SPACE(15)
@ 18,15 SAY "CODE CANNOT BE BLANK"
ELSE
USE MATER 
LOCATE FOR M_CODE=MC
IF FOUND()
@ 18,15 SAY "THIS CODE ALREADY EXISTS"
ELSE 

do while len(alltrim(mn))<=2
MN=SPACE(15)
@7,40 GET MN
READ
if len(alltrim(mn))<=2
@20,10 say "Please enter a valid name "
wait ""
@20,10 say "                          "
endif
enddo

OK="N"
OKok="N"
@18,15 SAY "IS DATA OK(Y/N) : "GET OK
READ
@18,15 SAY "R YOU SURE(Y/N) : "GET OKOK
READ
IF UPPER(OK)="Y" AND UPPER(OKOK)="Y"
APPEND BLANK
REPLACE M_CODE WITH MC
REPLACE M_NAME WITH MN
ENDIF
ENDIF
CLOSE ALL
ENDIF
@ 19,15 SAY "WANT TO CONTINUE (Y/N)? : "GET CH
READ
@ 18,15 SAY "                            "
ENDDO
DEAC WIND WIN1
RETURN


PROC CMO
ACTI WIND WIN1
USE MATER
BROW NOAPPEND NOEDIT NODELETE colo w+/bg
USE
DEAC WIND WIN1
rele all
RETURN

PROC CDE
ACTI WIND WIN1
USE MATER
LIST TO PRINT
CLOSE ALL
WAIT ""
CLEA
DEAC WIND WIN1
rele all
RETURN

PROC CSPW
OPW=SPACE(6)
NPW=SPACE(6)
NPW1=SPACE(6)
ACTI WIND WIN1
USE DO
GO TOP
STORE rtrim(PASS) TO PW
@10,10 SAY "ENTER OLD PASSWORD : " 
OPW=""
n=32
ctr=0
do while ctr<6
READ
@10,n say "*"
IF LASTKEY()<>13
	OPW=OPW+chr(LASTKEY())
	ctr=ctr+1
	n=n+1
ELSE
	CTR=6
ENDIF
enddo
IF OPW=PW
@11,10 SAY "ENTER NEW PASSWORD : "
NPW=""
n=32
ctr=0
do while ctr<6
READ
@11,n say "*"
IF LASTKEY()<>13
	NPW=NPW+chr(LASTKEY())
	ctr=ctr+1
	n=n+1
ELSE
	CTR=6
ENDIF
enddo
@12,10 SAY "NEW PASSWORD AGAIN : "
NPW1=""
n=32
ctr=0
do while ctr<6
READ
@12,n say "*"
IF LASTKEY()<>13
	NPW1=NPW1+chr(LASTKEY())
	ctr=ctr+1
	n=n+1
ELSE
	CTR=6
ENDIF
enddo
IF NPW=NPW1
OKOK="N"
OK="N"
@13,10 SAY "ARE YOU SURE(Y/N) : "GET OK
READ
@13,10 SAY "VERY SURE(Y/N)    : "GET OKOK
READ
IF UPPER(OK)="Y" AND UPPER(OKOK)="Y"
ZAP
APPEND BLANK
REPLACE PASS WITH NPW
ENDIF
ELSE
@12,10 SAY "TWO PASSWORDS DONT MATCH..TRY AGAIN.."
WAIT ""
ENDIF
ELSE
@12,10 SAY "YOU ARE NOT AUTHORIZED TO SET PASSWORDS.."
WAIT ""
ENDIF
CLOSE ALL
DEAC WIND WIN1
RETURN


PROC opPW
OPW=SPACE(6)
NPW=SPACE(6)
NPW1=SPACE(6)
ACTI WIND WIN1
pw=pwrd
@10,10 SAY "ENTER OLD PASSWORD : " 
OPW=""
n=32
ctr=0
do while ctr<6
READ
@10,n say "*"
IF LASTKEY()<>13
	OPW=OPW+chr(LASTKEY())
	ctr=ctr+1
	n=n+1
ELSE
	CTR=6
ENDIF
enddo
IF OPW=PW
@11,10 SAY "ENTER NEW PASSWORD : "
NPW=""
n=32
ctr=0
do while ctr<6
READ
@11,n say "*"
IF LASTKEY()<>13
	NPW=NPW+chr(LASTKEY())
	ctr=ctr+1
	n=n+1
ELSE
	CTR=6
ENDIF
enddo
@12,10 SAY "NEW PASSWORD AGAIN : "
NPW1=""
n=32
ctr=0
do while ctr<6
READ
@12,n say "*"
IF LASTKEY()<>13
	NPW1=NPW1+chr(LASTKEY())
	ctr=ctr+1
	n=n+1
ELSE
	CTR=6
ENDIF
enddo
IF NPW=NPW1
OKOK="N"
OK="N"
@13,10 SAY "ARE YOU SURE(Y/N) : "GET OK
READ
@13,10 SAY "VERY SURE(Y/N)    : "GET OKOK
READ
IF UPPER(OK)="Y" AND UPPER(OKOK)="Y"
use oper
locate for oname=onam
if found()
	REPLACE Pword WITH NPW
endif
use
ENDIF
ELSE
@12,10 SAY "TWO PASSWORDS DONT MATCH..TRY AGAIN.."
WAIT ""
ENDIF
ELSE
@12,10 SAY "YOU ARE NOT AUTHORIZED TO SET PASSWORDS.."
WAIT ""
ENDIF
DEAC WIND WIN1
RETURN


PROC CNEW
ACTI WIND WIN1
CH="y"
DO WHILE CH="y" OR CH="Y"
CLEA
@ 3,25 SAY "..... CUSTOMER ENTRY SCREEN ....."
CC=0
CN=SPACE(25)
AD=SPACE(25)
ON=SPACE(11)
OD=date()
OQ=0
OB=0
pbal=0
@ 6,20 SAY "CUSTOMER CODE    "
@ 7,20 SAY "CUSTOMER NAME    "
@ 8,20 SAY "CUSTOMER ADDRESS "
@ 9,20 SAY "ORDER NUMBER     "
@10,20 SAY "DATE             "
*@11,20 SAY "MATERIAL CODE    "
@11,20 SAY "QUANTITY         "
@11,55 say "TON"
@6,39 GET CC PICT '99999999'
READ
IF CC=0
@ 18,15 SAY "CODE CANNOT BE ZERO"
ELSE
USE CUST 
LOCATE FOR C_CODE=CC
IF FOUND()
@ 18,15 SAY "THIS CODE ALREADY EXISTS"
ELSE 

CN=SPACE(25)
do while len(alltrim(cn))<=7
CN=SPACE(25)
@7,40 GET CN
read
if len(alltrim(cn))<=7
@20,10 say "Please enter a valid name"
wait ""
@20,10 say "                          "
endif
enddo

AD=SPACE(25)
do while len(alltrim(ad))<=7
AD=SPACE(25)
@8,40 GET AD
read
if len(alltrim(ad))<=7
@20,10 say "Please enter a valid address"
wait ""
@20,10 say "                          "
endif
enddo

do while len(alltrim(on))<>11
@9,40 GET ON pict '99999999999'
read
if len(alltrim(on))<>11
@20,10 say "DO entry should be full"
wait ""
@20,10 say "                       "
endif
enddo

od={}
do while od={}
@10,40 GET OD
read
if od={}
@20,10 say "date cannot be blank"
wait ""
@20,10 say "                       "
endif
enddo

*matcod=0
*matnam=space(15)
*define popup matpop prompt fields m_name from 11,40
*use mater
*on key label enter deac popup matpop
*acti popup matpop
*if lastkey()=13
*	scatter memvar
*	matcod=m.m_code
*	matnam=m.m_name
*endif
*use
*on key label enter
*@11,40 say matcod
*@11,52 say matnam

oq=0
do while oq<=9
@11,40 GET OQ PICT '9999999.99'
READ
if oq<=9
@20,10 say "Please check order quantity"
wait ""
@20,10 say "                                        "
endif
enddo

oq=oq*1000
use cust
go top
locate for c_orderno=on
if found()
@ 18,15 SAY "DUPLICATE DO NUMBER NOT ALLOWED.."
else
OK="N"
OKOK="N"
@18,15 SAY "IS DATA OK(Y/N) : "GET OK
READ
@18,15 SAY "IS DATA OK(Y/N) : "GET OKOK
READ
IF UPPER(OK)="Y" AND UPPER(OKOK)="Y"
OB=OQ
APPEND BLANK
REPLACE C_CODE WITH CC
REPLACE C_NAME WITH CN
REPLACE C_ADDRESS WITH AD
REPLACE C_ORDERNO WITH ON
REPLACE O_DATE WITH OD
REPLACE O_QUANTITY WITH OQ
REPLACE O_BALANCE WITH OB
ENDIF
endif
ENDIF
CLOSE ALL
ENDIF
@ 19,15 SAY "WANT TO CONTINUE (Y/N)? : "GET CH
READ
@ 18,15 SAY "                            "
ENDDO
DEAC WIND WIN1
rele all
RETURN


PROC CADQ
ACTI WIND WIN1
PW=SPACE(6)
PWO=SPACE(6)
USE DO
STORE rtrim(PASS) TO PWO
CLOSE ALL
@2,30 SAY  "   ADD D.O. QUANTITY   " COLO W+/R*
@10,10 SAY "YOUR PASSWORD : "
PW=""
n=32
ctr=0
do while ctr<6
READ
@10,n say "*"
IF LASTKEY()<>13
	PW=PW+chr(LASTKEY())
	ctr=ctr+1
	n=n+1
ELSE
	CTR=6
ENDIF
enddo
IF PW=PWO
CH="y"
DO WHILE CH="y" OR CH="Y"
CLEA
@2,30 SAY  "   ADD D.O. QUANTITY   " COLO W+/R*
@ 3,25 SAY "..... CUSTOMER ENTRY SCREEN ....."
CC=0
CN=SPACE(25)
AD=SPACE(25)
ON=SPACE(11)
OD=CTOD("  /  /  ")
OQ=0
OB=0
pbal=0
@ 6,20 SAY "CUSTOMER CODE    "
@ 7,20 SAY "CUSTOMER NAME    "
@ 8,20 SAY "CUSTOMER ADDRESS "
@ 9,20 SAY "ORDER NUMBER     "
@10,20 SAY "DATE             "
@11,20 SAY "QUANTITY         "
@11,55 say "TON"
@6,39 GET CC PICT '99999999'
READ
USE CUST
GO TOP 
LOCATE FOR C_CODE=CC
IF FOUND()
STORE C_CODE TO CC
STORE C_NAME TO CN
@7,40 SAY CN
STORE C_ADDRESS TO AD
@8,40 SAY AD
STORE C_ORDERNO TO PON
do while len(alltrim(on))<>11
@9,40 GET ON pict '99999999999'
read
if len(alltrim(on))<>11
@20,10 say "DO entry should be full"
wait ""
@20,10 say "                       "
endif
enddo
STORE O_DATE TO POD
STORE O_BALANCE TO PBAL

go top
locate for c_orderno=on
if found()
	@19,15 say "Duplicate order number not allowed"
	wait ""
else
******
if pbal<10001
	go top
	LOCATE FOR C_CODE=CC
	@10,40 GET OD
	@11,40 GET OQ PICT '9999999.99'
	READ
	oq=oq*1000
	@12,20 SAY "PREV. BALANCE"
	@12,55 SAY "Kg"
	@12,40 SAY STR(PBAL,13,2) 
	OB=PBAL+OQ
	@13,20 SAY "FROM DO NO."
	@13,40 SAY PON
	@14,20 SAY "BALANCE"
	@14,55 SAY "Kg"
	@14,40 SAY STR(OB,13,2) 
	READ
	OK="N"
	OKOK="N"
	@18,15 SAY "IS DATA OK(Y/N) : "GET OK
	READ
	@18,15 SAY "R YOU SURE(Y/N) : "GET OKOK
	READ
	IF UPPER(OK)="Y" AND UPPER(OKOK)="Y"
		REPLACE C_ORDERNO WITH ON
		REPLACE O_DATE WITH OD 
		REPLACE O_QUANTITY WITH OQ 
		REPLACE P_BALANCE WITH PBAL
		REPLACE FROM_DO WITH PON
		REPLACE O_BALANCE WITH OB
	ENDIF
else
@20,15 SAY "DO CAN'T BE ADDED..PREVIOUS BALANCE EXCEEDS 10 TON"
endif
endif
ELSE
@19,15 SAY "NO SUCH CUSTOMER.."
WAIT ""
ENDIF
CLOSE ALL
@ 19,15 SAY "WANT TO CONTINUE (Y/N)? : "GET CH
READ
@ 18,15 SAY "                            "
ENDDO
ELSE
@12,10 SAY "YOU ARE NOT AUTHORIZED FOR THIS JOB.."
WAIT ""
ENDIF
DEAC WIND WIN1
rele all
RETURN


PROC CREQ
ACTI WIND WIN1
PW=SPACE(6)
PWO=SPACE(6)
USE DO
STORE rtrim(PASS) TO PWO
CLOSE ALL
@2,30 SAY  " REDUCE ORDER QUANTITY " COLO W+/R*
@10,10 SAY "YOUR PASSWORD : "
PW=""
n=32
ctr=0
do while ctr<6
READ
@10,n say "*"
IF LASTKEY()<>13
	PW=PW+chr(LASTKEY())
	ctr=ctr+1
	n=n+1
ELSE
	CTR=6
ENDIF
enddo
IF PW=PWO
CH="y"
DO WHILE CH="y" OR CH="Y"
CLEA
@2,30 SAY  " REDUCE ORDER QUANTITY " COLO W+/R*
@ 3,25 SAY "..... CUSTOMER ENTRY SCREEN ....."
CC=0
CN=SPACE(25)
AD=SPACE(25)
ON=SPACE(11)
OD=CTOD("  /  /  ")
OQ=0
OB=0
pbal=0
@ 6,20 SAY "CUSTOMER CODE    "
@ 7,20 SAY "CUSTOMER NAME    "
@ 8,20 SAY "CUSTOMER ADDRESS "
@10,20 SAY "DATE             "
@11,20 SAY "QUANTITY         "
@11,55 say "TON"
@6,39 GET CC PICT '99999999'
READ
USE CUST
GO TOP 
LOCATE FOR C_CODE=CC
IF FOUND()
STORE C_CODE TO CC
STORE C_NAME TO CN
@7,40 SAY CN
STORE C_ADDRESS TO AD
@8,40 SAY AD

STORE C_ORDERNO TO ON
STORE from_do TO pON
STORE O_DATE TO POD
STORE O_BALANCE TO PBAL
go top
	go top
	LOCATE FOR C_CODE=CC
	@10,40 GET OD
	@11,40 GET OQ PICT '9999999.99'
	READ
	oq=oq*1000
IF OQ<=PBAL
	@12,20 SAY "PREV. BALANCE"
	@12,55 SAY "Kg"
	@12,40 SAY STR(PBAL,13,2) 
	OB=PBAL-OQ
	@13,20 SAY "FROM DO NO."
	@13,40 SAY PON
	@14,20 SAY "BALANCE"
	@14,55 SAY "Kg"
	@14,40 SAY STR(OB,13,2) 
	READ
	OK="N"
	OKOK="N"
	@18,15 SAY "IS DATA OK(Y/N) : "GET OK
	READ
	@18,15 SAY "R YOU SURE(Y/N) : "GET OKOK
	READ
	IF UPPER(OK)="Y" AND UPPER(OKOK)="Y"
		REPLACE C_ORDERNO WITH ON
		REPLACE O_DATE WITH OD 
		REPLACE O_QUANTITY WITH OQ 
		REPLACE P_BALANCE WITH PBAL
		REPLACE FROM_DO WITH PON
		REPLACE O_BALANCE WITH OB
	ENDIF
ELSE
@20,15 SAY "QUANTITY TO REDUCE CAN'T BE MORE THAN EXISTING QUANTITY"  
ENDIF
ELSE
@19,15 SAY "NO SUCH CUSTOMER.."
WAIT ""
ENDIF
CLOSE ALL
@ 19,15 SAY "WANT TO CONTINUE (Y/N)? : "GET CH
READ
@ 18,15 SAY "                            "
ENDDO
ELSE
@12,10 SAY "YOU ARE NOT AUTHORIZED FOR THIS JOB.."
WAIT ""
ENDIF
DEAC WIND WIN1
rele all
RETURN

PROC CMOD
ACTI WIND WIN1
USE CUST
BROW NOEDIT NODELETE NOAPPEND colo w+/bg 
USE
DEAC WIND WIN1
rele all
RETURN
 
PROC CDP
ACTI WIND WIN1
set printer on
set devi to print
set printer to
use party
go top
store pname to ppname
store pinf to ppinf
store padd to ppadd
use
? CHR(18)+chr(14)+space(5)+ppname
? chr(14)+SPACE(5)+ppinf
? chr(14)+space(5)+ppadd
? CHR(14)+SPACE(5)+"       CUSTOMER REPORT"
? CHR(14)+SPACE(2)+"DATE: "+DTOC(DATE())+"    TIME: "+TIME()
USE
USE CUST
GO TOP
?chr(15)+"  CODE    NAME                       ADDRESS                    ORDER NO                   DATE            QUANTITY    BALANCE"  
?chr(15)+"------------------------------------------------------------------------------------------------------------------------------"  
do while not eof()
? chr(15)+str(c_code,8)+"  "+c_name+"  "+c_address+"  "+c_orderno+"  "+dtoc(o_date)+" "+str(o_quantity,13,2)+" "+str(o_balance,13,2)
skip
enddo
?
?
?
set printer to
set devi to scre
set printer off
?chr(18)
CLOSE ALL
WAIT " "
CLEA
DEAC WIND win1
rele all
RETURN


PROC CIND
ON KEY LABEL ESC DO EE
ACTI WIND WIN1
SELE 1 
USE CUST
acti wind wino
acti wind wino1
BROW noedit nodelete noappend colo w+/bg
deac wind wino1
deac wind wino

set printer on
set devi to print
set printer to
SELE 2
use party
go top
store pname to ppname
store pinf to ppinf
store padd to ppadd
USE
SELE 1
? CHR(18)+chr(14)+space(5)+ppname
? chr(14)+SPACE(5)+ppinf
? chr(14)+space(5)+ppadd
? CHR(14)+SPACE(2)+"     INDIVIDUAL CUSTOMER REPORT"
? CHR(14)+SPACE(2)+"DATE: "+DTOC(DATE())+"    TIME: "+TIME()
?
? chr(18)+"                  CODE            : "+str(c_code,8)
? chr(18)+"                  NAME            : "+c_name
? chr(18)+"                  ADDRESS         : "+c_address
? chr(18)+"                  ORDER NO        : "+c_orderno
? chr(18)+"                  DATE            : "+dtoc(o_date)
? chr(18)+"                  ORDER QUANTITY  : "+str(o_quantity,13,2)+" Kg"
? chr(18)+"                  PREVIOUS BALANCE: "+str(P_BALANCE,13,2)+" Kg"
? chr(18)+"                  FROM ORDER NO   : "+FROM_DO
? chr(18)+"                  BALANCE LEFT    : "+str(o_balance,13,2)+" Kg"
?
? chr(18)+"-------------------------------------------------------------"
?
?
set printer to
set devi to scre
set printer off
?chr(18)
CLOSE ALL
WAIT " "
CLEA
DEAC WIND win1
ON KEY LABEL ESC DO EEE
rele all
RETURN


PROC CAVF
ACTI WIND WIN1
USE CANE
go bottom
BROW NOEDIT NOAPPEND NODELETE colo w+/bg
USE 
DEAC WIND WIN1
rele all
RETURN

PROC SIVF
ACTI WIND WIN1
USE SIMPLE
go bottom
BROW NOAPPEND NOEDIT NODELETE colo w+/bg 
USE
DEAC WIND win1
rele all
RETURN

PROC SIFW
rele all
ACTI WIND WIN1
FW=0

***** READING COM PORT********************
acti wind win111
close all
if machine=1 
run wtread
b=fopen("weight.dat",10)
endif
if machine=2 
!basica ccc
b=fopen("cc.dbf",10)
endif
if machine=3 
clea
b=fopen("cc.dbf",10)
endif

a=fread(b,10)
clea
store val(a) to fW
CLOSE ALL
deac wind win111
*****************************************
if fw<21
CLEA
DO HEAD
SEA=SPACE(9)
SE=0
DASE=0
CS=SPACE(25)
CN=SPACE(25)
opn=space(20)
AD=SPACE(25)
MN=space(15)
VN=SPACE(15)
@ 6,16 SAY "YEAR         :"
@ 8,16 SAY "DAILY SERIAL :"
@ 9,16 SAY "DATE         :"
@10,16 SAY "TIME         :"
@11,16 SAY "OPERATOR NAME:"
@12,16 SAY "VEHICLE NO.  :"
@13,16 say "NAME         :"
@14,16 SAY "ADD/ DEST.   :"
@15,16 say "CHALLAN/SO.NO:"
@16,16 SAY "COLY/MAT     :"

IF MONTH(DATE())>3
STORE RIGHT(STR(YEAR(DATE())),4)+"-"+RIGHT(STR(YEAR(DATE())+1),4) TO SEA
ELSE
STORE RIGHT(STR(YEAR(DATE())-1),4)+"-"+RIGHT(STR(YEAR(DATE())),4) TO SEA
ENDIF
USE SIMPLE
GO BOTTOM
IF DATE_IN=DATE()
STORE (D_SERIAL+1) TO DASE
ELSE
STORE 1 TO DASE
ENDIF
close all
@ 6,31 SAY SEA
@ 8,31 SAY DASE PICT '99999'
@ 9,31 SAY DATE()
@10,31 SAY LEFT(TIME(),5)
opn=onam
@11,31 say OPN
@12,31 GET VN
@13,31 GET CN
@14,31 GET AD
@15,31 GET CS
@16,31 GET MN
read
if len(alltrim(opn))=0
opn="**********                 "
@11,31 say OPN
endif
if len(alltrim(vn))=0
vn="**********                 "
@12,31 say VN
endif
if len(alltrim(cn))=0
cn="**********                 "
@13,31 say CN
endif
if len(alltrim(ad))=0
ad="**********                 "
@14,31 say AD
endif
if len(alltrim(cs))=0
cs="**********                 "
@15,31 say CS
endif
if len(alltrim(mn))=0
mn="**********                 "
@16,31 say MN
endif
store "n" to blank
if opn="**********" or vn="**********" or cn="**********" or ad="**********" or cs="**********" or mn="**********"
store "y" to blank
endif
if blank="n"
@17,16 SAY "FIRST WEIGHT : Press any key to read"
WAIT ""

***** READING COM PORT********************
acti wind win111
close all
if machine=1 
run wtread
b=fopen("weight.dat",10)
endif
if machine=2
!basica ccc
b=fopen("cc.dbf",10)
endif
if machine=3
clea
b=fopen("cc.dbf",10)
endif

a=fread(b,10)
clea
store val(a) to fW
CLOSE ALL
deac wind win111
*****************************************
@17,31 say "                      "
@17,31 SAY FW PICT '9999999'+" Kg"
OK="y"
OKOK="N"
@ 21,20 SAY "IS DATA OK (Y/N)? "GET OK
READ
@ 21,20 SAY "R YOU SURE (Y/N)? "GET OKOK
READ
IF UPPER(OK)="Y" AND UPPER(OKOK)="Y"
USE SIMPLE
APPEND BLANK
REPLACE SEASON WITH SEA 
REPLACE D_SERIAL WITH DASE
REPLACE DATE_IN WITH DATE()
REPLACE TIME_IN WITH LEFT(TIME(),5)
REPLACE C_NAME WITH CN
REPLACE C_ADDRESS WITH AD
REPLACE CHALLAN_SO WITH CS
REPLACE MATERIAL WITH MN
REPLACE V_NO WITH VN
REPLACE O_NAME WITH OPN
REPLACE FIRST_WT WITH FW
CLOSE ALL
ppname=space(30)
ppinf=space(30)
ppadd=space(30)
use party
go top
store pname to ppname
store pinf to ppinf
store padd to ppadd
close all
CHO=SPACE(1)
@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
READ
IF CHO="P" OR CHO="p"
do while cho="p" or cho="P"
clea
set printer on
set device to print
set printer to
? chr(14)+space(5)+ppname
? chr(14)+SPACE(5)+ppinf
? chr(14)+SPACE(5)+ppadd
?
? chr(18)+space(27)+"YEAR    : "+sea
? chr(18)+SPACE(27)+"DATE   : "+DTOC(DATE())
? SPACE(27)+"TIME IN: "+LEFT(TIME(),5)
?
?
? SPACE(20)+"DAILY SERIAL : "+STR(DASE,6)
?
? SPACE(20)+"OPERATOR NAME: "+OPN
? SPACE(20)+"VEHICLE NO   : "+VN
? SPACE(20)+"NAME         : "+CN
? SPACE(20)+"ADDRESS      : "+AD
? SPACE(20)+"CHALLAN/SO.NO: "+CS
? SPACE(20)+"COLY/MAT NAME: "+MN 
if blank="y"
? space(20)+"Some fields are left blank..."
endif
?
? chr(14)+SPACE(5)+"FIRST WEIGHT :"+STR(FW,10)+" Kg" 
?
?
? chr(18)+"C.I.S.F.                                            Weighing Operator" 
? "---------------------------------------------------------------------"
?
?
set printer to
set device to screen
set printer off
cho=space(1)
@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
READ
enddo
WAIT " "
endif
ENDIF
else
@21,15 say "Some fields were blank...Weighment terminated.."
wait ""
endif
CLEA
DEAC WIND WIN1
DO HEAD
do tail
SET COLO TO W+/B
close all
rele all
else
@10,20 say "There is load on the platform"
@11,20 say "Make sure the platform is clear..."
wait ""
endif
deac wind win1
RETURN

PROC testing
rele all
ACTI WIND WIN1
CLEA
DO HEAD
CH="Y"
DO WHILE UPPER(CH)="Y"
@10,16 SAY "  T E S T I N G  "COLO W+/R
@17,16 SAY "READ WEIGHT   : Press any key to read"
WAIT ""

***** READING COM PORT********************
acti wind win111
close all
if machine=1 
run wtread
b=fopen("weight.dat",10)
endif
if machine=2 
!basica ccc
b=fopen("cc.dbf",10)
endif
if machine=3
clea
b=fopen("cc.dbf",10)
endif

a=fread(b,10)
clea
store val(a) to fW
CLOSE ALL
deac wind win111
*****************************************
@17,31 say "                      "
@17,31 SAY FW PICT '9999999'+" Kg"

choice="n"
@19,16 say "Print it(y/n) : "
@19,31 get choice
read
if upper(choice)="Y"
use party
go top
store pname to ppname
store pinf to ppinf
store padd to ppadd
use
set printer on
set devi to print
set printer to
? CHR(18)+chr(14)+space(5)+ppname
? chr(14)+SPACE(5)+ppinf
? chr(14)+space(5)+ppadd
? CHR(14)+SPACE(5)+"     TEST REPORT PRINT"
?
? chr(18)+space(20)+"date        : "+dtoc(date())
? chr(18)+space(20)+"Time        : "+time()
? chr(18)+space(20)+"Weight read : "+str(fw,10)+" Kg"
?
?CHR(15)+"-----------------------------------------------------------------------------------------------------------------------------------------"
?"N.B:"+chr(15)+"                                                            Designed & Installed by DADHWAL WEIGHING"
?CHR(15)+"-----------------------------------------------------------------------------------------------------------------------------------------"

set printer to
set devi to screen
set printer off
endif
wait ""
clea
@20,16 SAY "READ AGAIN(Y/N) ? : "GET CH
READ
ENDDO
DO HEAD
do tail
SET COLO TO W+/B
close all
rele all
deac wind win1
RETURN


PROC SISW
rele all
ACTI WIND WIN1
FW=0
***** READING COM PORT********************
acti wind win111
close all
if machine=1 
run wtread
b=fopen("weight.dat",10)
endif
if machine=2 
!basica ccc
b=fopen("cc.dbf",10)
endif
if machine=3
clea
b=fopen("cc.dbf",10)
endif

a=fread(b,10)
clea
store val(a) to fW
CLOSE ALL
deac wind win111
*****************************************
if fw<21

SEA=SPACE(9)
DASE=0
dain=ctod("  /  /  ")
daout=ctod("  /  /  ")
tin=space(5)
tout=space(5)
CN=SPACE(25)
vn=space(15)
AD=SPACE(25)
OPN=SPACE(20)
opn2=space(20)
MN=SPACE(15)
fw=0
sw=0
nw=0
STORE DATE() TO DAIN
@ 9,15 say "                DATE IN : "GET DAIN
READ
@ 11,15 SAY "SERIAL NUMBER TO RECALL : "GET daSE PICT '999999'
READ
SELE 1
USE SIMPLE
LOCATE FOR D_SERIAL=DASE AND DATE_IN=DAIN
IF FOUND()
	IF TIME_OUT<>"     "
	@13,15 SAY "DONT TRY AGAIN......SAVE COAL..SAVE NATION.."
	WAIT ""
ELSE
	store season to sea
	store d_serial to dase
	store date_in to dain
	store date() to daout
	store time_in to tin	
	store left(time(),5) to tout
	STORE C_NAME TO CN
	STORE C_ADDRESS TO AD
	STORE CHALLAN_SO TO CS
	STORE MATERIAL TO MN
	STORE O_NAME TO OPN
	STORE V_NO TO VN
	STORE FIRST_WT TO FW
	clea

	DO HEAD
	@ 4,18 say "YEAR   : "	
	@ 4,28 SAY SEA
	@ 5,18 say "DATE IN : "
	@ 5,28 SAY DAIN
	@5,38 SAY "DATE OUT : "
	@5,48 SAY DATE()	
	@ 8,22 SAY "DAILY SERIAL :"
	@ 9,22 SAY "TIME IN      :"
	@ 9,50 SAY "TIME OUT     :"
	@ 10,22 SAY "1st OPERATOR :"
	@ 11,22 SAY "2nd OPERATOR :"
	@ 12,22 SAY "VEHICLE NO.  :"
	@ 13,22 SAY "NAME         :"
	@ 14,22 SAY "ADDRESS      :"
	@ 15,22 SAY "CHALLAN/SO.NO:"
	@ 16,22 SAY "COLY/MAT NAME:"
	@ 17,22 SAY "FIRST WEIGHT :"
	@ 8,37 SAY DASE PICT '999999'
	@ 9,37 SAY TIN
	@9,65 SAY TOUT
	@10,37 SAY OPN
	@12,37 SAY VN
	@13,37 SAY CN
	@14,37 SAY AD
	@15,37 SAY CS
	@16,37 SAY MN
	opn2=onam
	@11,37 say OPN2
	if len(alltrim(opn2))=0
	opn2="**********                 "
	@11,37 say opn2
	endif
store "n" to blank
if opn="**********" or vn="**********" or cn="**********" or ad="**********" or cs="**********" or mn="**********" or opn2="**********"
store "y" to blank
endif
if blank="n"
	@17,37 SAY FW PICT '9999999'+" Kg"
	@ 18,22 SAY "SECOND WEIGHT: Press any key to read"
	wait ""

***** READING COM PORT********************
acti wind win111
if machine=1 
run wtread
b=fopen("weight.dat",10)
endif
if machine=2
!basica ccc
b=fopen("cc.dbf",10)
endif
if machine=3
clea
b=fopen("cc.dbf",10)
endif

a=fread(b,10)
clea
store val(a) to sW
deac wind win111
*****************************************
	if fw>sw 
	store (fw-sw) to nw
	else 
	store (sw-fw) to nw
	endif
	@ 19,22 SAY "NET WEIGHT   :"
	@ 18,37 say "                                 "
	@18,37 SAY SW PICT '9999999'+" Kg"
	@19,37 SAY NW PICT '9999999'+" Kg"

if nw<21
	@21,10 say "Net weight is too low..Zero weight taken"
	wait ""
	@21,10 say "                                        "
	sw=fw
	nw=0
endif

if nw<500 and nw<>0
	@21,10 say "Net weight is less than 500 kg..Try again"
	wait ""
else
	OK="n"
	OKOK="N"
	@21,20 SAY "IS DATA OK(Y/N) : "GET OK
	READ
	@21,20 SAY "R YOU SURE(Y/N) : "GET OKOK
	READ
	IF UPPER(OK)="Y" AND UPPER(OKOK)="Y"
		SELE 1
		REPLACE DATE_OUT WITH DATE()
		REPLACE TIME_OUT WITH LEFT(TIME(),5)
		REPLACE SECOND_WT WITH SW 
		replace o2_name with opn2
		CLOSE ALL
		ppname=space(30)
		ppinf=space(30)
		ppadd=space(30)
		use party
		go top
		store pname to ppname
		store pinf to ppinf
		store padd to ppadd
		close all
		CHO=SPACE(1)
		@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
		READ
		IF CHO="P" OR CHO="p"
			do while cho="P" or cho="p"
			clea
			set printer on
			set device to print
			set printer to
			? chr(14)+space(5)+ppname
			? chr(14)+SPACE(5)+ppinf
			? chr(14)+SPACE(5)+ppadd
			?
			? chr(18)+space(27)+"YEAR    : "+sea
			?
			? SPACE(25)+"DAILY SERIAL : "+STR(DASE,6)
			? 
			? SPACE(25)+"1st OPERATOR : "+OPN
			? SPACE(25)+"2nd OPERATOR : "+OPN2
			? SPACE(25)+"VEHICLE NO   : "+VN
			? SPACE(25)+"NAME         : "+CN
			? SPACE(25)+"ADDRESS      : "+AD
			? SPACE(25)+"CHALLAN/SO.NO: "+CS
			? SPACE(25)+"COLY/MAT     : "+MN
			if blank="y"
			? space(25)+"Some fields are left blank..."
			endif
			?
			? space(10)+"DATE IN      : "+DTOC(DAIN)+SPACE(5)+"DATE OUT     : "+DTOC(DATE())
			? SPACE(10)+"TIME IN      : "+TIN+SPACE(10)+"TIME OUT     : "+TOUT
			? chr(14)+SPACE(5)+"FIRST WEIGHT : "+STR(FW,10)+" Kg" 
			? chr(14)+SPACE(5)+"SECOND WEIGHT: "+STR(SW,10)+" Kg"
			? chr(14)+SPACE(5)+"NET WEIGHT   : "+STR(NW,10)+" Kg"
			?
			?
			? chr(18)+"C.I.S.F.                                            Weighing Operator" 
			? "-----------------------------------------------------------------------------"
			? "N.B:"+chr(15)+"Please check the weight before leaving.                                            Designed & Installed by DADHWAL WEIGHING"
			? "-----------------------------------------------------------------------------"
			? CHR(18)
			? 
			set printer to
			set device to screen
			set printer off
			CHO=SPACE(1)
			@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
			READ
			enddo
			WAIT " "
		ENDIF
	endif
ENDIF
else
@21,15 say "Some fields were left blank...Weighment terminated.."
wait ""
endif
****
ENDIF
ELSE
@ 13,15 SAY "SERIAL NUMBER NOT PRESENT"
WAIT ""
ENDIF
close all
CLEA
else
@10,20 say "There is load on the platform"
@11,20 say "Make sure the platform is clear..."
wait ""
endif
DEAC WIND WIN1
DO HEAD
do tail
SET COLO TO W+/B
rele all
RETURN


PROC SREDW
ACTI WIND WIN1
STORE DATE() TO DA1
STORE DATE() TO DA2
@ 12,15 SAY "DATE FROM : "GET DA1
@ 13,15 SAY "DATE TO   : "GET DA2
READ
SELE 1
USE SIMPLE
SET FILTER TO DATE_OUT>=DA1 AND DATE_OUT<=DA2
LN=0
SET PRINTER ON
SET DEVI TO PRINT
SET PRINTER TO
GO TOP
TOT=0
DO WHILE NOT EOF()
SELE 1
STORE SEASON TO SEA
STORE D_SERIAL TO SL
STORE DATE_IN TO DIN
STORE DATE_OUT TO DOUT
STORE TIME_IN TO TIN
STORE TIME_OUT TO TOUT
STORE V_NO TO VN
STORE C_NAME TO CN
STORE C_ADDRESS TO AD
STORE CHALLAN_SO TO CS
STORE Material TO MN
STORE FIRST_WT TO FW
STORE SECOND_WT TO SW
IF FW>SW
STORE (FW-SW) TO NW
ELSE 
STORE (SW-FW) TO NW
ENDIF
IF LN=0
WAIT ""
DO HED2
ENDIF
TOT=TOT+NW
?CHR(15)+"|"+STR(SL,6)+" |"+TIN+" |"+TOUT+" |"+dtoc(dout)+" |"+CN+" |"+vn+" |"+MN+" |"+STR(FW,10)+" |"+STR(SW,10)+" |"+STR(NW,10)+" |"
?CHR(15)+"|"+SPACE(6)+" |"+SPACE(5)+" |"+SPACE(5)+" |"+space(10)+" |"+AD+" |"+space(15)+" |"+SPACE(15)+" |"+SPACE(10)+" |"+SPACE(10)+" |"+SPACE(10)+" |"
SELE 1
LN=LN+1
IF LN>=26
LN=0
ENDIF
SKIP
ENDDO
?CHR(15)+"------------------------------------------------------------------------------------------------------------------------------------"
?CHR(18)+"                                         TOTAL OF NET WEIGHT = "+STR(TOT,10)+" Kg"
?CHR(15)+"------------------------------------------------------------------------------------------------------------------------------------"
? 
?
SET PRINTER TO
SET DEVI TO SCRE
SET PRINTER OFF
SET FILTER TO
CLOSE ALL
DEAC WIND win1
rele all
RETURN


PROC SREOW
on key label esc do ee
ACTI WIND WIN1
USE SIMPLE
acti wind wino
acti wind wino1
go bottom
BROW noedit nodelete noappend colo w+/bg
deac wind wino1
deac wind wino
STORE SEASON TO SEA
STORE D_SERIAL TO DASE
STORE DATE_IN TO DIN
STORE DATE_OUT TO DOUT
STORE TIME_IN TO TIN
STORE TIME_OUT TO TOUT
STORE V_NO TO VN
STORE C_NAME TO CN
STORE C_ADDRESS TO AD
STORE CHALLAN_SO TO CS
STORE Material TO MN
STORE FIRST_WT TO FW
STORE SECOND_WT TO SW
IF FW>SW
STORE (FW-SW) TO NW
ELSE 
STORE (SW-FW) TO NW
ENDIF
CLOSE ALL

USE PARTY 
GO TOP
store pname to ppname
store pinf to ppinf
store padd to ppadd
close all

CHO=SPACE(1)
@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
READ
IF CHO="P" OR CHO="p"
do while cho="P" or cho="p"
clea
set printer on
set device to print
set printer to
? chr(14)+space(5)+ppname
? chr(14)+SPACE(5)+ppinf
? chr(14)+SPACE(5)+ppadd
?
? chr(18)+SPACE(27)+"DATE    : "+DTOC(dout)
?
? SPACE(25)+"DAILY SERIAL : "+STR(DASE,6)
? SPACE(25)+"TIME IN      : "+TIN
? SPACE(25)+"TIME OUT     : "+TOUT
? 
? SPACE(25)+"VEHICLE NO   : "+VN
? SPACE(25)+"NAME         : "+CN
? SPACE(25)+"ADDRESS      : "+AD
? SPACE(25)+"COLY/MAT     : "+MN
?
? SPACE(25)+"FIRST WEIGHT : "+STR(FW,10)+" Kg" 
? SPACE(25)+"SECOND WEIGHT: "+STR(SW,10)+" Kg"
? SPACE(25)+"NET WEIGHT   : "+STR(NW,10)+" Kg"
?
?
? chr(18)+"C.I.S.F.                                            Weighing Operator" 
? "-----------------------------------------------------------------------------"
? "N.B:"+chr(15)+"Please check the weight before leaving.                                            Designed & Installed by DADHWAL WEIGHING"
? CHR(18)+"-----------------------------------------------------------------------------"
IF NW=0
? "              EMPTY TRUCK OUT, ANY WEIGHT, IF TAKEN IGNORED"
ENDIF
?
set printer to
set device to screen
set printer off
CHO=SPACE(1)
@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
READ
enddo
WAIT " "
ENDIF
WAIT ""
CLEA
DEAC WIND WIN1
DO HEAD
do tail
on key label esc do eee
rele all
RETURN


PROC HED2
ppname=space(30)
ppinf=space(30)
ppadd=space(30)
SELE 4
use party
go top
store pname to ppname
store pinf to ppinf
store padd to ppadd
USE
? CHR(18)+chr(14)+space(5)+ppname
? chr(14)+SPACE(5)+ppinf
? CHR(14)+SPACE(5)+ppadd
?
? CHR(18)+"FROM DATE : "+DTOC(DA1)+"        TO DATE : "+DTOC(DA2)
?CHR(15)+"------------------------------------------------------------------------------------------------------------------------------------"
?CHR(15)+"|TR.NO. | T.IN |T.OUT |  DATEOUT  |  CUSTOMER NAME           |  VEHICLE NO.   | COLY/MAT NAME  | FIRST WT. | SECOND WT.| NET WT.   |"
?CHR(15)+"------------------------------------------------------------------------------------------------------------------------------------"
rele all
RETURN

***********************************************************************
***************** S P E C I A L     W E I G H I N G *******************
***********************************************************************

PROC COFW
rele all
ACTI WIND WIN1
fw=0
OPN=SPACE(20)

***** READING COM PORT********************
acti wind win111
close all
if machine=1 
run wtread
b=fopen("weight.dat",10)
endif
if machine=2
!basica ccc
b=fopen("cc.dbf",10)
endif
if machine=3
clea
b=fopen("cc.dbf",10)
endif

a=fread(b,10)
clea
store val(a) to fW
CLOSE ALL
deac wind win111
*****************************************
if fw<21
CLEA
DO HEAD
@  6,6 SAY "YEAR          "
@  7,6 SAY "DAILY SERIAL  "
@  8,6 SAY "DATE          "
@ 8,41 SAY "TIME          "
@  9,6 SAY "VEHICLE NO.   "
@ 9,41 SAY "OPERATOR NAME "
@ 10,6 say "PARTY CODE    "
@ 11,6 SAY "PARTY NAME    "
@ 12,6 SAY "PARTY ADDRESS "
@ 13,6 SAY "ORDER NO.     "
@ 14,6 SAY "ORDER DATE    "
@ 15,6 SAY "COLY/MAT CODE "
@ 16,6 SAY "COLY/MAT NAME "
iF MONTH(DATE())>3
STORE RIGHT(STR(YEAR(DATE())),4)+"-"+RIGHT(STR(YEAR(DATE())+1),4) TO SEA
ELSE
STORE RIGHT(STR(YEAR(DATE())-1),4)+"-"+RIGHT(STR(YEAR(DATE())),4) TO SEA
ENDIF
SELE 1
USE DWI
GO BOTTOM
IF DATE_IN=DATE()
STORE (SL_NO+1) TO SNO
ELSE
STORE 1 TO SNO
ENDIF
@ 6,21 SAY SEA
@ 7,21 SAY SNO PICT '99999'
@ 8,21 SAY DATE()
@ 8,56 SAY LEFT(TIME(),5)
VN=SPACE(15)
TC=0
mc=space(8)
@ 9,21 GET VN
opn=onam
@ 9,56 say OPN
@ 10,21 GET TC pict '99999999'
READ
if len(alltrim(vn))=0
vn="**********          "
@ 9,21 say VN
endif
if len(alltrim(opn))=0
opn="**********         "
@ 9,56 say OPN
endif
SELE 2
use cust
LOCATE FOR C_CODE=TC
IF FOUND()
store c_name to CN
store c_address to ad
STORE C_ORDERNO TO ON
STORE O_DATE TO OD
store o_balance to obac
@11,21 SAY CN
@11,55 say " Balance Qnty  " color w+/r*
@12,55 say str(obac/1000,10)+" Ton " color w+/r*
@12,21 SAY AD
@13,21 SAY ON
@14,21 SAY OD
@ 15,21 get MC
read
SELE 3
USE MATER
go top
LOCATE FOR UPPER(M_CODE)=UPPER(MC)
IF FOUND()
STORE M_NAME TO MN
@ 16,21 SAY MN
store "n" to blank
if opn="**********" or vn="**********" 
store "y" to blank
endif
if blank="n"
@18,6 SAY "FIRST WEIGHT  : Press any key to continue "
wait ""
@ 18,19 say "                             "

***** READING COM PORT********************
acti wind win111
if machine=1 
run wtread
b=fopen("weight.dat",10)
endif
if machine=2
!basica ccc
b=fopen("cc.dbf",10)
endif
if machine=3
clea
b=fopen("cc.dbf",10)
endif

a=fread(b,10)
clea
store val(a) to fW
deac wind win111
*****************************************
if fw>=5000
	@ 18,21 SAY FW PICT '9999999'+" Kg"
	OK="y"
	OKOK="N"
	@ 21,15 SAY "IS DATA OK (Y/N)? "GET OK
	READ
	@ 21,15 SAY "R YOU SURE (Y/N)? "GET OKOK
	READ
	IF UPPER(OK)="Y" AND UPPER(OKOK)="Y"
		SELE 1
		APPEND BLANK
		REPLACE SEASON WITH SEA 
		REPLACE SL_NO WITH SNO
		REPLACE DATE_IN WITH DATE()
		REPLACE TIME_IN WITH LEFT(TIME(),5)
		REPLACE TC_CODE WITH tC
		REPLACE O_NAME WITH OPN
		REPLACE V_NO WITH VN
		REPLACE TM_CODe WITH MC
		REPLACE FIRST_WT WITH FW
		CLOSE ALL
		ppname=space(30)
		ppinf=space(30)
		ppadd=space(30)
		sele 4
		use party
		go top
		store pname to ppname
		store pinf to ppinf
		store padd to ppadd
		use
		CHO=SPACE(1)
		@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
		READ
		do while cho="P" or cho="p"
		if CHO="P" OR CHO="p"
			clea
			set printer on
			set device to print
			set printer to
			? chr(18)+chr(14)+space(5)+ppname
			? chr(14)+SPACE(5)+ppinf
			? chr(14)+space(5)+ppadd
			?
			? chr(18)+SPACE(17)+"YEAR   : "+SEA+SPACE(1)+"DATE   : "+DTOC(DATE())
			? SPACE(27)+"TIME   : "+LEFT(TIME(),5)
			?
			? SPACE(20)+"SERIAL       : "+STR(SNO,6)
			? SPACE(20)+"CODE         : "+STR(tC,8)
			? SPACE(20)+"OPERATOR NAME: "+OPN
			? SPACE(20)+"VEHICLE NO   : "+VN
			? SPACE(20)+"PARTY NAME   : "+CN
			? SPACE(20)+"ADDRESS      : "+AD
			? SPACE(20)+"ORDER NO     : "+ON
			? SPACE(20)+"ORDER DATE   : "+DTOC(OD)
			? SPACE(20)+"COLY/MAT CODE: "+MC
			? SPACE(20)+"COLY/MAT     : "+MN
			if blank="y"
				? space(20)+"Some fields are left blank"
			endif
			?
			? CHR(14)+SPACE(5)+"FIRST WEIGHT :"+STR(FW,10)+" Kg" 
			?
			? CHR(18)+"C.I.S.F.                                            Weighing Operator" 
			?
			?
			set printer to
			set device to screen
			set printer off
		ENDIF
		CHO=SPACE(1)
		@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
		READ
		WAIT ""
		enddo
	ENDIF
ELSE
@21,15 say "First weight cannot be less than 5 ton"
wait ""
ENDIF
else
@21,15 say "Some fields were left blank..Weighment terminated.."
wait ""
endif
ELSE 
@ 21,10 SAY "     INVALID CODE           "
WAIT " "
ENDIF
ELSE
@ 21,10 SAY "THIS CUSTOMER DOES NOT EXIST"
WAIT " "
endif
DEAC WIND WIN1
CLOSE ALL
DO HEAD
do tail
rele all
else
@10,20 say "There is load on the platform"
@11,20 say "Make sure the platform is clear..."
wait ""
deac wind win1
endif
RETURN

PROC COSW
rele all
ACTI WIND WIN1
FW=0

***** READING COM PORT********************
acti wind win111
close all
if machine=1 
run wtread
b=fopen("weight.dat",10)
endif
if machine=2
!basica ccc
b=fopen("cc.dbf",10)
endif
if machine=3
clea
b=fopen("cc.dbf",10)
endif

a=fread(b,10)
clea
store val(a) to fW
CLOSE ALL
deac wind win111
*****************************************
if fw<21
TC=0
SEA=SPACE(9)
SN=0
DASE=0
dain=ctod("  /  /  ")
daout=ctod("  /  /  ")
tin=space(5)
tout=space(5)
vn=space(15)
AD=SPACE(15)
OPN=SPACE(20)
OPN2=SPACE(20)
VC=SPACE(7)
cv=space(15)
cr=0
fw=0
sw=0
nw=0
bm=0
fiw=0
STORE DATE() TO DAIN
@ 9,15 say "                DATE IN : "GET DAIN
READ
@ 11,15 SAY "SERIAL NUMBER TO RECALL : "GET SN PICT '999999'
READ
SELE 1
USE DWI
LOCATE FOR sl_NO=SN and date_in=dain
IF FOUND()
IF TIME_OUT<>"     "
@13,15 SAY "DONT TRY AGAIN...SAVE COAL..SAVE NATION.."
WAIT ""
ELSE
store season to sea
store SL_NO to dase
store date_in to dain
store date() to daout
store time_in to tin
store left(time(),5) to tout
store tC_code to tc
store first_wt to fw
STORE V_NO TO VN
STORE O_NAME TO OPN
STORE TM_CODE TO VC
STORE FIRST_WT TO FW
CLEA
DO HEAD
@ 4,18 say "YEAR   : "
@ 4,28 SAY SEA
@ 5,18 say "DATE IN : "
@ 5,28 SAY DAIN
@5,38 SAY "DATE OUT : "
@5,48 SAY DATE()
@ 7,1 SAY "SERIAL       :"
@ 8,1 SAY "TIME IN      :"
@ 9,1 SAY "TIME OUT     :"
@ 10,1 SAY "1st OPERATOR :"
@ 11,1 SAY "2nd OPERATOR :"
@ 12,1 SAY "VEHICLE NO   :"
@ 13,1 SAY "CODE         :"
@ 14,1 SAY "PARTY NAME   :"
@ 15,1 SAY "ADDRESS      :"
@ 16,1 SAY "ORDER NO.    :"
@ 17,1 SAY "ORDER DATE   :"
@ 18,1 SAY "COLY/MAT     :"
@ 7,46 SAY "FIRST WEIGHT(Kg) :"
SELE 2 
USE CUST 
LOCATE FOR C_CODE=TC
store c_name to cN
store c_address to ad
STORE C_ORDERNO TO ON
STORE O_DATE TO OD 
STORE O_BALANCE TO OBA
SELE 3
USE MATER
GO TOP
LOCATE FOR upper(M_CODE)=upper(VC)
IF FOUND()
STORE M_NAME TO CV
endif
@ 7,16 SAY DASE PICT '999999'
@ 8,16 SAY TIN
@9,16 SAY TOUT 
@10,16 SAY OPN
@12,16 SAY VN
@13,16 SAY TC
@14,16 SAY CN
@15,16 SAY AD
@16,16 SAY ON
@17,16 SAY OD
@18,16 SAY CV 
opn2=onam
@11,16 say OPN2
	if len(alltrim(opn2))=0
	opn2="**********                 "
	@11,16 say opn2
	endif
store "n" to blank
if opn="**********" or vn="**********" or opn2="**********"
store "y" to blank
endif
if blank="n"
@7,65 SAY FW PICT '9999999'+" Kg"
@ 8,46 SAY "SECOND WEIGHT(Kg):Press any key"
WAIT ""
@ 8,64 SAY "             "
***** READING COM PORT********************
acti wind win111
if machine=1 
run wtread
b=fopen("weight.dat",10)
endif
if machine=2
!basica ccc
b=fopen("cc.dbf",10)
endif
if machine=3
clea
b=fopen("cc.dbf",10)
endif

a=fread(b,10)
clea
store val(a) to sW
deac wind win111
*****************************************

if fw>sw 
store (fw-sw) to nw
else 
store (sw-fw) to nw
endif

if nw<21
	@21,10 say "Net weight is too low..Zero weight taken"
	sw=fw
	nw=0
	@9,46 SAY "NET WEIGHT(Kg)   :"
	@8,65 SAY SW PICT '9999999'+" Kg"
	@9,65 SAY "No weight"
	wait ""
	@21,10 say "                                        "
endif

IF OBA-NW>=0
	@9,46 SAY "NET WEIGHT(Kg)   :"
	@8,65 SAY SW PICT '9999999'+" Kg"
	@9,65 SAY NW PICT '9999999'+" Kg"

if nw<9000 and nw<>0 
	@21,10 say "Net weight cannot be less than 9 ton"
	@9,46 SAY "NET WEIGHT(Kg)   :"
	@8,65 SAY "Too low     "
	@9,65 SAY "Wrong weight"
else
	@9,46 SAY "NET WEIGHT(Kg)   :"
	@8,65 SAY SW PICT '9999999'+" Kg"
	@9,65 SAY NW PICT '9999999'+" Kg"
	OK="N"
	OKOK="N"
	@22,20 SAY "IS DATA OK(Y/N) : "GET OK
	READ
	@22,20 SAY "R YOU SURE(Y/N) : "GET OKOK
	READ
	IF UPPER(OK)="Y" AND UPPER(OKOK)="Y"
		SELE 2
		REPLACE O_BALANCE WITH (OBA-NW)
		STORE NW*CR TO AMT
		SELE 1 
		REPLACE DATE_OUT WITH DATE()
		REPLACE TIME_OUT WITH LEFT(TIME(),5)
		REPLACE SECOND_WT WITH SW 
		REPLACE o2_name WITH opn2 
		CLOSE ALL
		ppname=space(30)
		ppinf=space(30)
		ppadd=space(30)
		use party
		go top
		store pname to ppname
		store pinf to ppinf
		store padd to ppadd
		use
		CHO=SPACE(1)
		@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
		READ
		do while cho="p" or cho="P"
		IF CHO="P" OR CHO="p"
			clea
			set printer on
			set device to print
			set printer to
			? chr(18)+chr(14)+space(5)+ppname
			? chr(14)+SPACE(5)+ppinf
			? chr(14)+space(5)+ppadd
			?
			? chr(18)+SPACE(17)+"YEAR   : "+SEA
			?
			? chr(18)+SPACE(1)+"DAILY SERIAL : "+STR(DASE,6)+SPACE(15)+"CODE         : "+STR(TC,8)
			? SPACE(1)+"DATE IN      : "+DTOC(DAIN)+SPACE(11)+"DATE OUT     : "+DTOC(DAOUT)
			? SPACE(1)+"TIME IN      : "+TIN+SPACE(16)+"PARTY NAME   : "+CN
			? SPACE(1)+"TIME OUT     : "+TOUT+SPACE(16)+"ADDRESS      : "+AD
			? SPACE(1)+"1st OPERATOR : "+OPN+"    2nd OPERATOR : "+OPN2
			? SPACE(1)+"VEHICLE NO   : "+VN+SPACE(6)+"ORDER NO     : "+ON
			? SPACE(1)+"COLY/MAT     : "+CV+SPACE(6)+"ORDER DATE   : "+DTOC(OD)
			if blank="y"
			?space(1)+"Some fields are left blank"
			endif
			?
			? chr(14)+SPACE(5)+"FIRST WEIGHT(Kg) :"+STR(FW,10)+" Kg" 
			? CHR(14)+SPACE(5)+"SECOND WEIGHT(Kg):"+STR(SW,10)+" Kg"
			? chr(14)+SPACE(5)+"NET WEIGHT(Kg)   :"+STR(NW,10)+" Kg"
			?
			? "C.I.S.F                                             Weighing Operator" 
			?CHR(15)+"-----------------------------------------------------------------------------------------------------------------------------------------"
			?"N.B:"+chr(15)+"Check the weight before leaving.                                                            Designed & Installed by DADHWAL WEIGHING"
			?CHR(15)+"-----------------------------------------------------------------------------------------------------------------------------------------"
			? CHR(18)
			?
			set printer to
			set device to screen
			set printer off
		ENDIF
		CHO=SPACE(1)
		@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
		READ
		wait ""
		enddo
	ENDIF
endif
else 
@ 21,15 say "BALANCE NOT AVAILABLE, WEIGHMENT TERMINATED......"
ENDIF
else
@21,15 say "Some fields were left blank...Weighment terminated.."
wait ""
endif
ENDIF
ELSE
@ 13,15 SAY "SERIAL NUMBER NOT PRESENT"
ENDIF

WAIT ""
CLEA
DEAC WIND WIN1
CLOSE ALL
DO HEAD
do tail
rele all
else
@10,20 say "There is load on the platform"
@11,20 say "Make sure the platform is clear..."
wait ""
deac wind win1
endif
RETURN


PROC COVF
acti wind win1
use dwi
GO BOTTOM
brow noedit nodelete noappend colo w+/bg
use
deac wind win1
rele all
RETURN 


PROC REDW
ACTI WIND WIN1
STORE DATE() TO DA1
STORE DATE() TO DA2
@ 12,15 SAY "DATE FROM : "GET DA1
@ 13,15 SAY "DATE TO   : "GET DA2
READ
SELE 1
use dwi
SET FILTER TO DATE_OUT>=DA1 AND DATE_OUT<=DA2
LN=0
SET PRINTER ON
SET DEVI TO PRINT
SET PRINTER TO
GO TOP
tot=0
DO WHILE NOT EOF()
STORE SEASON TO SEA
STORE SL_NO TO SL
STORE DATE_IN TO DIN
STORE DATE_OUT TO DOUT
STORE TIME_IN TO TIN
STORE TIME_OUT TO TOUT
STORE TC_CODE TO CC
STORE V_NO TO VN
STORE TM_CODE TO MC
STORE FIRST_WT TO FW
STORE SECOND_WT TO SW
IF FW>SW
STORE (FW-SW) TO NW
ELSE 
STORE (SW-FW) TO NW
ENDIF
tot=tot+nw
SELE 2
USE CUST
GO TOP
LOCATE FOR C_CODE=CC
CN=SPACE(25)
AD=SPACE(25)
ON=SPACE(25)
OD=CTOD("  /  /  ")
IF FOUND()
STORE C_NAME TO CN
STORE C_ADDRESS TO AD
STORE C_ORDERNO TO ON
STORE O_DATE TO OD
ENDIF
SELE 3
USE MATER
GO TOP
LOCATE FOR M_CODE=MC
MN=SPACE(15)
MR=0
IF FOUND()
STORE M_NAME TO MN
ENDIF
IF LN=0
WAIT ""
DO HEAD1 with "  Date wise Weighment Report"

ENDIF
?CHR(15)+"       |"+STR(SL,5)+" |"+DTOC(Din)+","+tin+" |"+CN+" |"+ON+" |"+VN+" |"+STR(FW,10)+" |"+STR(SW,10)+" |"+STR(NW,10)+" |"
?CHR(15)+"       |"+SPACE(5)+" |"+DTOC(Dout)+","+tout+" |"+AD+" |"+DTOC(OD)+SPACE(1)+" |"+MN+" |"+SPACE(10)+" |"+SPACE(10)+" |"+SPACE(10)+" |"
SELE 1
LN=LN+1
IF LN>=26
LN=0
ENDIF
SKIP
ENDDO
?CHR(15)+"       -----------------------------------------------------------------------------------------------------------------------"
?chr(18)+"                                     Total of net weight= "+str(tot,10)+" Kg" 
?CHR(15)+"       -----------------------------------------------------------------------------------------------------------------------"
? 
?CHR(15)+"       Signature of ASM                                                                                  Signature of Operator"
?
SET PRINTER TO
SET DEVI TO SCRE
SET PRINTER OFF
SET FILTER TO
CLOSE ALL
DEAC WIND win1
rele all
RETURN


PROC RESM
ACTI WIND WIN1
STORE DATE() TO DA1
STORE DATE() TO DA2
TMCODE=SPACE(8)
@ 10,15 SAY "COLY/MAT CODE : "GET TMCODE
@ 12,15 SAY "DATE FROM     : "GET DA1
@ 13,15 SAY "DATE TO       : "GET DA2
READ
SELE 1
use dwi
SET FILTER TO DATE_OUT>=DA1 AND DATE_OUT<=DA2 AND TM_CODE=TMCODE
LN=0
SET PRINTER ON
SET DEVI TO PRINT
SET PRINTER TO
GO TOP
tot=0
DO WHILE NOT EOF()
STORE SEASON TO SEA
STORE SL_NO TO SL
STORE DATE_IN TO DIN
STORE DATE_OUT TO DOUT
STORE TIME_IN TO TIN
STORE TIME_OUT TO TOUT
STORE TC_CODE TO CC
STORE V_NO TO VN
STORE TM_CODE TO MC
STORE FIRST_WT TO FW
STORE SECOND_WT TO SW
IF FW>SW
STORE (FW-SW) TO NW
ELSE 
STORE (SW-FW) TO NW
ENDIF
tot=tot+nw
SELE 2
USE CUST
GO TOP
LOCATE FOR C_CODE=CC
CN=SPACE(25)
AD=SPACE(25)
ON=SPACE(25)
OD=CTOD("  /  /  ")
IF FOUND()
STORE C_NAME TO CN
STORE C_ADDRESS TO AD
STORE C_ORDERNO TO ON
STORE O_DATE TO OD
ENDIF
SELE 3
USE MATER
GO TOP
LOCATE FOR M_CODE=MC
MN=SPACE(15)
MR=0
IF FOUND()
STORE M_NAME TO MN
ENDIF
IF LN=0
WAIT ""
DO HEAD1  with "Material Name: "+mn
ENDIF
?CHR(15)+"       |"+STR(SL,5)+" |"+DTOC(Din)+","+tin+" |"+CN+" |"+ON+" |"+VN+" |"+STR(FW,10)+" |"+STR(SW,10)+" |"+STR(NW,10)+" |"
?CHR(15)+"       |"+SPACE(5)+" |"+DTOC(Dout)+","+tout+" |"+AD+" |"+DTOC(OD)+SPACE(1)+" |"+MN+" |"+SPACE(10)+" |"+SPACE(10)+" |"+SPACE(10)+" |"
SELE 1
LN=LN+1
IF LN>=26
LN=0
ENDIF
SKIP
ENDDO
?CHR(15)+"       -----------------------------------------------------------------------------------------------------------------------"
?chr(18)+"                                     Total of net weight= "+str(tot,10)+" Kg" 
?CHR(15)+"       -----------------------------------------------------------------------------------------------------------------------"
? 
?CHR(15)+"       Signature of ASM                                                                                  Signature of Operator"
?
SET PRINTER TO
SET DEVI TO SCRE
SET PRINTER OFF
SET FILTER TO
CLOSE ALL
DEAC WIND win1
rele all
RETURN


PROC RESP
ACTI WIND WIN1
STORE DATE() TO DA1
STORE DATE() TO DA2
TPCODE=0
@ 10,15 SAY "PARTY CODE : "GET TPCODE PICT '99999999'
@ 12,15 SAY "DATE FROM  : "GET DA1
@ 13,15 SAY "DATE TO    : "GET DA2
READ
SELE 1
use dwi
go top
SET FILTER TO DATE_OUT>=DA1 AND DATE_OUT<=DA2 AND TC_CODE=TPCODE
LN=0
SET PRINTER ON
SET DEVI TO PRINT
SET PRINTER TO
GO TOP
tot=0
DO WHILE NOT EOF()
STORE SEASON TO SEA
STORE SL_NO TO SL
STORE DATE_IN TO DIN
STORE DATE_OUT TO DOUT
STORE TIME_IN TO TIN
STORE TIME_OUT TO TOUT
STORE TC_CODE TO CC
STORE V_NO TO VN
STORE TM_CODE TO MC
STORE FIRST_WT TO FW
STORE SECOND_WT TO SW
IF FW>SW
STORE (FW-SW) TO NW
ELSE 
STORE (SW-FW) TO NW
ENDIF
tot=tot+nw
SELE 2
USE CUST
GO TOP
LOCATE FOR C_CODE=CC
CN=SPACE(25)
AD=SPACE(25)
ON=SPACE(25)
OD=CTOD("  /  /  ")
IF FOUND()
STORE C_NAME TO CN
STORE C_ADDRESS TO AD
STORE C_ORDERNO TO ON
STORE O_DATE TO OD
ENDIF
SELE 3
USE MATER
GO TOP
LOCATE FOR M_CODE=MC
MN=SPACE(15)
MR=0
IF FOUND()
STORE M_NAME TO MN
ENDIF
IF LN=0
WAIT ""
DO HEAD1  with " Partyname: "+cn
ENDIF
?CHR(15)+"       |"+STR(SL,5)+" |"+DTOC(Din)+","+tin+" |"+CN+" |"+ON+" |"+VN+" |"+STR(FW,10)+" |"+STR(SW,10)+" |"+STR(NW,10)+" |"
?CHR(15)+"       |"+SPACE(5)+" |"+DTOC(Dout)+","+tout+" |"+AD+" |"+DTOC(OD)+SPACE(1)+" |"+MN+" |"+SPACE(10)+" |"+SPACE(10)+" |"+SPACE(10)+" |"
SELE 1
LN=LN+1
IF LN>=26
LN=0
ENDIF
SKIP
ENDDO
?CHR(15)+"       -----------------------------------------------------------------------------------------------------------------------"
?chr(18)+"                                     Total of net weight= "+str(tot,10)+" Kg" 
?CHR(15)+"       -----------------------------------------------------------------------------------------------------------------------"
? 
?CHR(15)+"       Signature of ASM                                                                                  Signature of Operator"
?
SET PRINTER TO
SET DEVI TO SCRE
SET PRINTER OFF
SET FILTER TO
CLOSE ALL
DEAC WIND win1
rele all
RETURN


PROC REMW
ACTI WIND WIN1
mn1=space(15)
STORE DATE() TO DA1
STORE DATE() TO DA2
@ 12,15 SAY "DATE FROM : "GET DA1
@ 13,15 SAY "DATE TO   : "GET DA2
READ

SELE 1
use dwi
INDEX ON TM_CODE TO DWIM
USE DWI INDEX DWIM

SET FILTER TO DATE_OUT>=DA1 AND DATE_OUT<=DA2
LN=0
SET PRINTER ON
SET DEVI TO PRINT
SET PRINTER TO
GO TOP
tot=0
TOT1=0
STORE TM_CODE TO MC1
DO WHILE NOT EOF()
STORE SEASON TO SEA
STORE SL_NO TO SL
STORE DATE_IN TO DIN
STORE DATE_OUT TO DOUT
STORE TIME_IN TO TIN
STORE TIME_OUT TO TOUT
STORE TC_CODE TO CC
STORE V_NO TO VN
STORE TM_CODE TO MC
STORE FIRST_WT TO FW
STORE SECOND_WT TO SW

IF FW>SW
STORE (FW-SW) TO NW
ELSE 
STORE (SW-FW) TO NW
ENDIF

tot=tot+nw
SELE 2
USE CUST
GO TOP
LOCATE FOR C_CODE=CC
CN=SPACE(25)
AD=SPACE(25)
ON=SPACE(25)
OD=CTOD("  /  /  ")

IF FOUND()
STORE C_NAME TO CN
STORE C_ADDRESS TO AD
STORE C_ORDERNO TO ON
STORE O_DATE TO OD
ENDIF

SELE 3
USE MATER
GO TOP
LOCATE FOR M_CODE=MC
MN=SPACE(15)
MR=0

IF FOUND()
STORE M_NAME TO MN
ENDIF

IF LN=0
WAIT ""
DO HEAD1  with "   Sorted Material Report"
ENDIF

IF MC<>MC1
?CHR(18)+SPACE(20)+"TOTAL OF "+MN1+" = "+STR(TOT1,10)+" Kg"
?CHR(15)
TOT1=0
LN=LN+1
ENDIF

TOT1=TOT1+NW
?CHR(15)+"       |"+STR(SL,5)+" |"+DTOC(Din)+","+tin+" |"+CN+" |"+ON+" |"+VN+" |"+STR(FW,10)+" |"+STR(SW,10)+" |"+STR(NW,10)+" |"
?CHR(15)+"       |"+SPACE(5)+" |"+DTOC(Dout)+","+tout+" |"+AD+" |"+DTOC(OD)+SPACE(1)+" |"+MN+" |"+SPACE(10)+" |"+SPACE(10)+" |"+SPACE(10)+" |"
SELE 1
LN=LN+1

IF LN>=26
LN=0
ENDIF

STORE TM_CODE TO MC1
SELE 3
STORE M_NAME TO MN1
SELE 1
SKIP
ENDDO
?CHR(18)+SPACE(20)+"TOTAL OF "+MN1+" = "+STR(TOT1,10)+" Kg"
?CHR(15)
?CHR(15)+"       -----------------------------------------------------------------------------------------------------------------------"
?chr(18)+"                                     Total of net weight= "+str(tot,10)+" Kg" 
?CHR(15)+"       -----------------------------------------------------------------------------------------------------------------------"
? 
?CHR(15)+"       Signature of ASM                                                                                  Signature of Operator"
?
SET PRINTER TO
SET DEVI TO SCRE
SET PRINTER OFF
SET FILTER TO
CLOSE ALL
DEAC WIND win1
rele all
RETURN


PROC REOS
opname=space(20)
totin=0
totout=0
ACTI WIND WIN1
STORE DATE() TO DA1
@ 12,15 SAY "DATE : "GET DA1
READ
da2=da1
SET PRINTER ON
SET DEVI TO PRINT
SET PRINTER TO
do heados
sele 1
use logs
sele 2
use dwi
sele 1
set filter to dat=da1
go top
totalin=0
totalout=0
do while not eof()
store oname to opname
store login to intime
store logout to outtime
sele 2
count to totin for date_in=da1 and o_name=opname
count to totout for date_in=da1 and o2_name=opname
? chr(14)+space(4)+"|"+opname+" |  "+intime+"  |  "+outtime+"  | "+str(totin,5)+" | "+str(totout,5)+" |"
totalin=totalin+totin
totalout=totalout+totout
sele 1
skip
enddo
close all
? chr(14)+space(30)+"Total in  :"+str(totalin)
? chr(14)+space(30)+"Total Out :"+str(totalout)
?CHR(15)+"        -----------------------------------------------------------------------------------------------------------------------"
?
?CHR(15)+"          Signature of ASM                                                                              Signature of Operator"
?CHR(15)+"        -----------------------------------------------------------------------------------------------------------------------"
?
SET PRINTER TO
SET DEVI TO SCRE
SET PRINTER OFF
DEAC WIND win1
rele all
RETURN


PROC REPW
ACTI WIND WIN1
mn1=space(15)
STORE DATE() TO DA1
STORE DATE() TO DA2
@ 12,15 SAY "DATE FROM : "GET DA1
@ 13,15 SAY "DATE TO   : "GET DA2
READ
SELE 1
use dwi
INDEX ON TC_CODE TO DWIM
USE DWI INDEX DWIM
SET FILTER TO DATE_OUT>=DA1 AND DATE_OUT<=DA2
LN=0
SET PRINTER ON
SET DEVI TO PRINT
SET PRINTER TO
GO TOP
tot=0
TOT1=0
STORE TC_CODE TO CC1
DO WHILE NOT EOF()
STORE SEASON TO SEA
STORE SL_NO TO SL
STORE DATE_IN TO DIN
STORE DATE_OUT TO DOUT
STORE TIME_IN TO TIN
STORE TIME_OUT TO TOUT
STORE TC_CODE TO CC
STORE V_NO TO VN
STORE TM_CODE TO MC
STORE FIRST_WT TO FW
STORE SECOND_WT TO SW
IF FW>SW
STORE (FW-SW) TO NW
ELSE 
STORE (SW-FW) TO NW
ENDIF
tot=tot+nw
SELE 2
USE CUST
GO TOP
LOCATE FOR C_CODE=CC
CN=SPACE(25)
AD=SPACE(25)
ON=SPACE(25)
OD=CTOD("  /  /  ")
IF FOUND()
STORE C_NAME TO CN
STORE C_ADDRESS TO AD
STORE C_ORDERNO TO ON
STORE O_DATE TO OD
ENDIF
SELE 3
USE MATER
GO TOP
LOCATE FOR M_CODE=MC
MN=SPACE(15)
MR=0
IF FOUND()
STORE M_NAME TO MN
ENDIF
IF LN=0
WAIT ""
DO HEAD1  with "     Sorted Party Report"
ENDIF

IF CC<>CC1
?CHR(18)+SPACE(20)+"TOTAL OF "+MN1+" = "+STR(TOT1,10)+" Kg"
?CHR(15)
TOT1=0
LN=LN+1
ENDIF
TOT1=TOT1+NW
?CHR(15)+"       |"+STR(SL,5)+" |"+DTOC(Din)+","+tin+" |"+CN+" |"+ON+" |"+VN+" |"+STR(FW,10)+" |"+STR(SW,10)+" |"+STR(NW,10)+" |"
?CHR(15)+"       |"+SPACE(5)+" |"+DTOC(Dout)+","+tout+" |"+AD+" |"+DTOC(OD)+SPACE(1)+" |"+MN+" |"+SPACE(10)+" |"+SPACE(10)+" |"+SPACE(10)+" |"

SELE 1
LN=LN+1
IF LN>=26
LN=0
ENDIF
STORE TC_CODE TO CC1
SELE 2
STORE C_NAME TO MN1
SELE 1
SKIP
ENDDO

?CHR(18)+SPACE(20)+"TOTAL OF "+MN1+" = "+STR(TOT1,10)+" Kg"
?CHR(15)
?CHR(15)+"       -----------------------------------------------------------------------------------------------------------------------"
?chr(18)+"                                     Total of net weight= "+str(tot,10)+" Kg" 
?CHR(15)+"       -----------------------------------------------------------------------------------------------------------------------"
? 
?CHR(15)+"       Signature of ASM                                                                                  Signature of Operator"
?
SET PRINTER TO
SET DEVI TO SCRE
SET PRINTER OFF
SET FILTER TO
CLOSE ALL
DEAC WIND win1
rele all
RETURN


PROC HEAD1 
parameters rtyp
ppname=space(30)
ppinf=space(30)
ppadd=space(30)
sele 4
use party
go top
store pname to ppname
store pinf to ppinf
store padd to ppadd
use
? CHR(18)+chr(14)+space(5)+ppname
? chr(14)+SPACE(5)+ppinf
? chr(14)+space(5)+ppadd
? CHR(14)+SPACE(5)+rtyp
?
? CHR(18)+"      FROM DATE : "+DTOC(DA1)+"         TO DATE : "+DTOC(DA2)
?CHR(15)+"       -----------------------------------------------------------------------------------------------------------------------"
?CHR(15)+"       |S.NO. |DATE,TIME IN/OUT | CUSTOMER NAME/ADDRESS    | O.NO./DATE | V.NO./ COL.MAT.| FIRST WT. | SECOND WT.|  NET WT.  |"
?CHR(15)+"       -----------------------------------------------------------------------------------------------------------------------"
rele all
RETURN


PROC HEADOS
ppname=space(30)
ppinf=space(30)
ppadd=space(30)
sele 4
use party
go top
store pname to ppname
store pinf to ppinf
store padd to ppadd
use
? CHR(18)+chr(14)+space(5)+ppname
? chr(14)+SPACE(5)+ppinf
? chr(14)+space(5)+ppadd
? CHR(14)+SPACE(5)+"Daily Operator Summary"
?
? CHR(18)+"      DATE : "+DTOC(DA1)
?CHR(15)+"        ---------------------------------------------------------------------------------------------------------------------"
?CHR(14)+space(3)+" | Operator            | Time in | Timeout |  In   |  Out  |"
?CHR(15)+"        ---------------------------------------------------------------------------------------------------------------------"
rele all
RETURN


PROC HEAD22
ppname=space(30)
ppinf=space(30)
ppadd=space(30)
SELE 4
use party
go top
store pname to ppname
store pinf to ppinf
store padd to ppadd
USE
? CHR(18)+chr(14)+space(5)+ppname
? chr(14)+SPACE(5)+ppinf
? CHR(14)+SPACE(5)+ppadd
?
? CHR(14)+"DATE : "+DTOC(DA1)+"         "+"SHIFT : "+SHI
?CHR(15)+"------------------------------------------------------------------------------------------------------------------------------------"
?CHR(15)+"|TR.NO. | T.IN |T.OUT |  DATE IN  |  CUSTOMER NAME           |  VEHICLE NO.   | MATERIAL NAME  | FIRST WT. | SECOND WT.| NET WT.   |"
?CHR(15)+"------------------------------------------------------------------------------------------------------------------------------------"
RETURN


PROC HEAD2
ppname=space(30)
ppinf=space(30)
ppadd=space(30)
SELE 4
use party
go top
store pname to ppname
store pinf to ppinf
store padd to ppadd
USE
? CHR(18)+chr(14)+space(5)+ppname
? chr(14)+SPACE(5)+ppinf
? CHR(14)+SPACE(5)+ppadd
?
? CHR(18)+"DATE FROM : "+DTOC(DA1)+"   "+"DATE TO : "+DTOC(DA2)
?CHR(15)+"------------------------------------------------------------------------------------------------------------------------------------"
?CHR(15)+"|TR.NO. | T.IN |T.OUT |  DATE IN  |  CUSTOMER NAME           |  VEHICLE NO.   | MATERIAL NAME  | FIRST WT. | SECOND WT.| NET WT.   |"
?CHR(15)+"------------------------------------------------------------------------------------------------------------------------------------"
RETURN

PROC HEAD
ppname=space(30)
ppinf=space(30)
ppadd=space(30)
sele 4
use party
go top
store pname to ppname
store pinf to ppinf
store padd to ppadd
SET COLO TO W+/n
@ 1,20 TO 5,59 CLEA
SET COLO TO W+/R
@ 0,19 TO 4,58 CLEA
@ 1,22 say ppname COLO W+/R*
@ 2,22 say ppinf
@ 3,22 say ppadd
use
SET COLO TO W+/B 
rele all
RETURN


proc tail
SET COLO TO W+/n 
@18,31 to 20,47 clea
SET COLO TO W+/r 
@17,30 to 19,46 clea
@18,34 say "  D W S"
SET COLO TO W+/b 
return

proc party
DEFINE WIND WIN11 FROM 17,0 TO 23,79 
ppname=space(30)
ppinf=space(30)
ppadd=space(30)
acti wind win11
@ 3,15 say "     enter password : "
@ 3,37 say " "
READ
@ 3,38 say "*"
PWD11=lastKEY()
READ
@ 3,39 say "*"
PWD12=lastKEY()
READ
@ 3,40 say "*"
PWD13=lastKEY()
READ
@ 3,41 say "*"
PWD14=lastKEY()
READ
@ 3,42 say "*"
PWD15=lastKEY()
READ
@ 3,43 say "*"
PWD16=lastKEY()

if (pwd11=100 and (pwd12=119 and (pwd13=105 and (pwd14=49 and (pwd15=50 and pwd16=51)))))
@ 2,15 say "enter party name    : "
@ 3,15 say "enter party nature  : "
@ 4,15 say "enter party address : "
@ 2,38 get ppname
@ 3,38 get ppinf
@ 4,38 get ppadd
read
if ppname<>space(30)
use party
zap
append blank
replace pname with ppname
replace pinf with ppinf
replace padd with ppadd
use
clea
ch11="N"
@ 2,15 say "Want to empty all the data files (Y/N) : "get ch11
read
if ch11="Y" or ch11="y"
	use simple	
	zap
	use
	
	use dwi
	zap
	use

	use cust
	zap
	use
	
	use mater
	zap
	use
endif	
endif
ELSE
CLEA
@ 2,20 SAY "YOU ARE NOT AUTHORISED FOR THIS JOB"
@ 3,20 SAY "Press any key to continue ........."
WAIT "" 
endif
deac wind win11
do head
do tail
rele all
return


PROC REOW
on key label esc do ee
ACTI WIND WIN1
SELE 1
USE DWI
acti wind wino
acti wind wino1
go bottom
brow noedit nodelete noappend colo w+/bg
deac wind wino1
deac wind wino
store season to sea
store SL_NO to dase
store date_in to dain
store date_out to daout
store time_in to tin
store time_out to tout
store tC_code to tc
store o_name to opn
store o2_name to opn2
store first_wt to fw
STORE V_NO TO VN
STORE TM_CODE TO VC
STORE FIRST_WT TO FW
store second_wt to sw
IF FW>SW
STORE (FW-SW) TO NW
ELSE 
STORE (SW-FW) TO NW
ENDIF 

SELE 2 
USE CUST 
LOCATE FOR C_CODE=TC
store c_name to cN
store c_address to ad
STORE C_ORDERNO TO ON
STORE O_DATE TO OD 

SELE 3
USE MATER
GO TOP
LOCATE FOR upper(M_CODE)=upper(VC)
STORE M_NAME TO CV
	use party
	go top
	store pname to ppname
	store pinf to ppinf
	store padd to ppadd
close all
CHO=SPACE(1)
@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
READ
do while cho="p" or cho="P"
IF CHO="P" OR CHO="p"
clea
set printer on
set device to print
set printer to
? chr(18)+chr(14)+space(5)+ppname
? chr(14)+SPACE(5)+ppinf
? chr(14)+space(5)+ppadd
? chr(18)+SPACE(17)+"(DUPLICATE WEIGH SLIP)"
? chr(18)+SPACE(17)+"YEAR   : "+SEA+SPACE(1)
?
? chr(18)+SPACE(1)+"DAILY SERIAL : "+STR(DASE,6)+SPACE(15)+"CODE         : "+STR(TC,8)
? SPACE(1)+"DATE IN      : "+DTOC(DAIN)+SPACE(11)+"DATE OUT     : "+DTOC(DAOUT)
? SPACE(1)+"TIME IN      : "+TIN+SPACE(16)+"PARTY NAME   : "+CN
? SPACE(1)+"TIME OUT     : "+TOUT+SPACE(16)+"ADDRESS      : "+AD
? SPACE(1)+"1st OPERATOR : "+OPN+"    2nd OPERATOR : "+OPN2
? SPACE(1)+"VEHICLE NO   : "+VN+SPACE(6)+"ORDER NO     : "+ON
? SPACE(1)+"COLY/MAT     : "+CV+SPACE(6)+"ORDER DATE   : "+DTOC(OD)
?
? chr(14)+SPACE(5)+"FIRST WEIGHT(Kg) :"+STR(FW,10)+" Kg" 
? CHR(14)+SPACE(5)+"SECOND WEIGHT(Kg):"+STR(SW,10)+" Kg"
? chr(14)+SPACE(5)+"NET WEIGHT(Kg)   :"+STR(NW,10)+" Kg"
?
? "C.I.S.F                                             Weighing Operator" 
? CHR(15)+"-----------------------------------------------------------------------------------------------------------------------------------------"
? "N.B:"+chr(15)+"Check the weight before leaving.                                                            Designed & Installed by DADHWAL WEIGHING"
? CHR(15)+"-----------------------------------------------------------------------------------------------------------------------------------------"
? CHR(18)
?
set printer to
set device to screen
set printer off
endif
CHO=SPACE(1)
@21,15 SAY "PRESS 'P' TO PRINT AND ANY OTHER KEY TO RETURN"GET CHO
READ
wait ""
enddo
DEAC WIND WIN1
CLOSE ALL
DO HEAD
do tail
on key label esc do eee
rele all
RETURN


proc admin
set sysmenu on
opch=0
acti wind win1
do while opch<>4
clea
@2,5 say "Please note that operators can change their passwords themselves"
@3,5 say "by pressing shift+f5 after logging in with their username password"
@4,5 say "The operators may be informed by the Manager, if required so"

@10,10 say "1. Manager Settings"
@11,10 say "2. User Settings"
@12,10 say "3. Machine Settings"
@13,10 say "4. Quit"
@14,10 say "Enter Choice : "
@14,26 get opch pict '9'
read
if opch=1 
use oper
set filter to oname="manager"
brow fields uname, pword noappend title "PRESS CTRL+C TO SAVE"
use oper
close all
endif
if opch=2 
use oper
set filter to oname<>"manager"
brow title "PRESS CTRL+N TO ADD USER, CTRL+T TO REMOVE USER, CTRL+C TO SAVE"
use oper
PACK
close all
endif
if opch=3
clea
mcch=0
@7,10 say  "Please select the machine"
@8,10 say  "1. Suzuki Machine"
@9,10 say  "2. Basic Machine"
@10,10 say  "3. Alpine Machine"
@12,10 say "Enter Choice: "get mcch pict '9'
read
if mcch=1 then
use mach
go top
replace mach with "Suzuki"
use
endif
if mcch=2 then
use mach
go top
replace mach with "Basic"
use
endif
if mcch=3 then
use mach
go top
replace mach with "Alpine"
use
endif

endif
enddo
deac wind all
clea
return


PROC QUIT
use logs
go top
locate for oname=onam and dat=date()
if found()
replace logout with left(time(),5)
endif
CLOSE ALL
CLEA
set clock off
rele all
set mouse on
deac wind all
DEAC MENU MAIN 
DEAC POPUP ALL
return     

return
