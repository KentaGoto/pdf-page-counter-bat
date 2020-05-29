@echo off

setlocal

set script_dir=%~dp0

rem PDF�̃t�@�C�����ƃy�[�W�����������܂��e�L�X�g�t�@�C��
set result_file=result.csv
set result=%script_dir%%result_file%

rem pdfinfo.exe
rem set pdfinfo_cmd=%script_dir%pdfinfo.exe
set pdfinfo_cmd=pdfinfo.exe

rem grep.exe
rem set grep_cmd=%script_dir%grep.exe
set grep_cmd=grep.exe

rem sed.exe
rem set sed_cmd=%script_dir%sed.exe
set sed_cmd=sed.exe

rem result.txt�������������
if exist %result% (
  del %result%
)

rem �Ώۃf�B���N�g��
set /p dir="Dir: "
cd "%dir%"

echo.

rem Header
echo File Name,Page Count >> %result%

rem �J�E���^�[
set counter=1

for /r %%i in (*.pdf) do (
  rem �������t�@�C�����R�}���h�v�����v�g�ɏo��
  echo %%i

  rem �t�@�C������result.csv�ɏo�́i���s���Ȃ��j
  set /P filename=""%%i""< NUL >> %result%

  rem �y�[�W����result.csv�ɏo��
  %pdfinfo_cmd% "%%i" | %grep_cmd% Pages: | %sed_cmd% -e "s/Pages: */,/" >> %result%

  set /a counter=counter+1
)

rem ��s�󂯂�
echo. >> %result%
rem �y�[�W���v���BExcel�̊֐��ō��v�����o���B
echo Total,=SUM(B2:B%counter%) >> %result%

echo.
echo Done^!
echo You can open the result in a result.csv file by pressing the Enter key.

pause > nul

call %result%

endlocal