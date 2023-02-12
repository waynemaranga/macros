@echo off
Title Creating a Winword shortcut with arguments on the desktop
Mode con cols=70 lines=5 & color 9E
REM Determine if the OS is (32/64 bits) to set the correct path of Program files.
IF /I "%PROCESSOR_ARCHITECTURE%"=="x86" (
        Set "strProgramFiles=%ProgramFiles%"
    ) else (
        Set "strProgramFiles=%programfiles(x86)%"
)

set Key="HKEY_CLASSES_ROOT\Word.Application\CurVer"
For /f "tokens=4 delims= " %%a in ('reg Query %Key% /ve ^| findstr /R "[0-9]"') do (
	for /f "tokens=3 delims=." %%b in ('echo %%a') do (
		Set "Ver=%%b"
	)
)

Rem The shortcut name with the .lnk extension
Set "MyShortcutName=%userprofile%\desktop\%~n0.lnk"
set "TargetPath=%strProgramFiles%\Microsoft Office\Office%ver%\Winword.EXE"
Rem Here we put the arguments of the command line
Set "Arguments=/q /a"

If not exist "%MyShortcutName%" (
	Call :CreateShortcut
) else (
	Goto Main
)
Exit

::***********************************************************************
:CreateShortcut
echo(
echo         Creating the shortcut on the desktop is in progress .....
::***********************************************************************
Powershell ^
$s=(New-Object -COM WScript.Shell).CreateShortcut('%MyShortcutName%'); ^
$s.TargetPath='"%TargetPath%"'; ^
$s.Arguments='%Arguments%'; ^
$s.Save()
Exit /b
::***********************************************************************
:Main
echo Hello
pause