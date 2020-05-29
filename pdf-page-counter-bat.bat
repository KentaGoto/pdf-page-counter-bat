@echo off

setlocal

set script_dir=%~dp0

rem PDFのファイル名とページ数が書き込まれるテキストファイル
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

rem result.txtがあったら消す
if exist %result% (
  del %result%
)

rem 対象ディレクトリ
set /p dir="Dir: "
cd "%dir%"

echo.

rem Header
echo File Name,Page Count >> %result%

rem カウンター
set counter=1

for /r %%i in (*.pdf) do (
  rem 処理中ファイルをコマンドプロンプトに出力
  echo %%i

  rem ファイル名をresult.csvに出力（改行しない）
  set /P filename=""%%i""< NUL >> %result%

  rem ページ数をresult.csvに出力
  %pdfinfo_cmd% "%%i" | %grep_cmd% Pages: | %sed_cmd% -e "s/Pages: */,/" >> %result%

  set /a counter=counter+1
)

rem 一行空ける
echo. >> %result%
rem ページ合計数。Excelの関数で合計数を出す。
echo Total,=SUM(B2:B%counter%) >> %result%

echo.
echo Done^!
echo You can open the result in a result.csv file by pressing the Enter key.

pause > nul

call %result%

endlocal