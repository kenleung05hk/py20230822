@if (@CodeSection == @Batch) @then
@echo off
taskkill /im "EXCEL.EXE"
timeout /t 1
SET SendKeys=CScript //nologo //E:JScript "%~F0"

start "lmeprice" "Z:\Dealing Room\ken leung\py data\LME closing price\LME closing price.xlsx"
timeout /t 19
%SendKeys% "%1"
timeout /t 2
taskkill /im "EXCEL.EXE"
timeout /t 4
%SendKeys% "%S"


@end
var WshShell = WScript.CreateObject("WScript.Shell");
WshShell.SendKeys(WScript.Arguments(0));