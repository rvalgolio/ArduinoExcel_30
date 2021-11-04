
rem put all the following files in a folder and run the batch shown below
rem	ArduinoExcel.dll
rem	ArduinoExcel.tlb
rem	Arduino_Excel_30.ino
rem	rExcel.h
rem	rExcel.cpp
rem	keywords.txt
rem	Arduino_Excel_30.xls
rem
rem install dll
rem if 64bit seems to be necessary to register twice the dll (either as 64 and 32 bit)
:CHECK_OS
IF EXIST "%PROGRAMFILES(X86)%" (GOTO 64BIT) ELSE (GOTO 32BIT)
:64BIT
Copy ArduinoExcel.dll C:\Windows\SysWOW64\
Copy ArduinoExcel.tlb C:\Windows\SysWOW64\
C:\Windows\Microsoft.Net\Framework64\v4.0.30319\RegAsm.exe C:\Windows\SysWOW64\ArduinoExcel.dll /unregister 
C:\Windows\Microsoft.Net\Framework64\v4.0.30319\RegAsm.exe C:\Windows\SysWOW64\ArduinoExcel.dll /tlb:C:\Windows\SysWOW64\ArduinoExcel.tlb /codebase
:32BIT
Copy ArduinoExcel.dll C:\Windows\System32\
Copy ArduinoExcel.tlb C:\Windows\System32\
C:\Windows\Microsoft.Net\Framework\v4.0.30319\RegAsm.exe C:\Windows\System32\ArduinoExcel.dll /unregister
C:\Windows\Microsoft.Net\Framework\v4.0.30319\RegAsm.exe C:\Windows\System32\ArduinoExcel.dll /tlb:C:\Windows\System32\ArduinoExcel.tlb /codebase

rem copy Arduino Sketch
IF EXIST "%userprofile%\Documents\Arduino\Arduino_Excel_30\Arduino_Excel_30.ino" (GOTO LIBRARY_DIR)
mkdir "%userprofile%\Documents\Arduino\Arduino_Excel_30"
copy Arduino_Excel_30.ino "%userprofile%\Documents\Arduino\Arduino_Excel_30"

rem copy rExcel library
:LIBRARY_DIR
IF EXIST "%userprofile%\Documents\Arduino\libraries\rExcel" (GOTO LIBRARY_FILES)
mkdir "%userprofile%\Documents\Arduino\libraries\rExcel"
rem copy files anyway (useful for updates)
:LIBRARY_FILES
copy rExcel.h "%userprofile%\Documents\Arduino\libraries\rExcel"
copy rExcel.cpp "%userprofile%\Documents\Arduino\libraries\rExcel"
copy keywords.txt "%userprofile%\Documents\Arduino\libraries\rExcel"

rem copy Excel file and instructions
:EXCEL
rem IF EXIST "%userprofile%\Documents\Arduino_Excel\Arduino_Excel_30.xls" (GOTO INSTRUCTIONS)
mkdir "%userprofile%\Documents\Arduino_Excel"
copy Arduino_Excel_30.xls "%userprofile%\Documents\Arduino_Excel"

rem open instructions
:INSTRUCTIONS
copy Instructions.pdf "%userprofile%\Documents\Arduino_Excel"
"C:\Program Files\Internet Explorer\iexplore.exe" "%userprofile%\Documents\Arduino_Excel\Instructions.pdf"
