Bat_To_Exe_Converter_x64.exe /bat "ExtractAll_en-US.bat" ^
/exe "Release/ExtractAll_en-US.exe" ^
/x64 ^
/include "7z.exe" ^
/include "unrar.exe" ^
/workdir "%~dp0" ^
/extractdir "%~dp0" ^
/deleteonexit ^
/overwrite ^
/attributes ^
/display
