set PYTHON=c:\Python27\python.exe

set DATAFILE=resources/data3.xls
set VISITNO=T1V1
set VISITYEAR=2012
set VISITMONTH=1

%PYTHON% code\generate.py %DATAFILE% %VISITNO% %VISITYEAR% %VISITMONTH%
%PYTHON% code\convert_to_pdf.py 

