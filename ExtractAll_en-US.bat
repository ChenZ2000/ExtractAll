@echo off
setlocal EnableDelayedExpansion

:: Initialize variables
set "i=1"
set "DecompressFail=0"
set "DelFail=0"

echo copyright info ↓↓↓
echo UnRar.exe © win.rar GmbH All rights reserved. ^<https://www.win-rar.com/^>
echo 7z.exe © Igor Pavlov. ^<https://www.7-zip.org/^>
echo Script written and packaged by ChenZ ^<https://chenz.cloud^>
echo.
echo Welcome to ExtractAll^^!
:: Prompt the user to enter passwords
:PasswordLoop
set /p Password%i%="Enter password %i% (leave blank to go to the next step): "
if "!Password%i%!"=="" (
    set /a PasswordCount=!i!-1
    goto AskExtractOption
) else (
    set /a i+=1
    goto PasswordLoop
)

:AskExtractOption
echo.
:: Ask if each archive should be extracted into its own subfolder
echo Would you like to extract each archive to its own subfolder?
set /p SeparateFolders=Please enter Y or N and press Enter: 
if /i "!SeparateFolders!"=="Y" (
    set "ExtractToSeparateFolders=1"
) else (
    set "ExtractToSeparateFolders=0"
)
echo.

:: Ask if the source archive should be deleted after extraction
echo Would you like to delete the archive after extraction?
set /p DeleteSource=Please enter Y or N and press Enter: 
if /i "!DeleteSource!"=="Y" (
    set "DeleteAfterFinish=1"
) else (
    set "DeleteAfterFinish=0"
)

echo.
echo Starting extraction...
echo.

:: Create the output directory (if it doesn't already exist)
if not exist "Output\" (
    mkdir "Output\"
    if !ERRORLEVEL! NEQ 0 (
        echo Failed to create output directory "Output\". Exiting.
        exit /b 1
    )
)

:: Process all archive files in the current directory
for %%F in (*.rar *.zip *.7z *.tar *.gz *.bz2 *.xz *.tar.gz *.tgz) do (
    echo.
    echo ----------------------------------------
    echo Processing file: %%F
    set "process=1"
    set "baseName=%%~nF"
    set "extension=%%~xF"

    echo Initial Base name: "!baseName!"
    echo Extension: "!extension!"

    :: Initialize variables for multipart archive handling
    set "isMultipart=0"
    set "isFirstPart=0"
    set "newBaseName=!baseName!"

    :: For RAR files, check for multipart archives with .partN pattern (underscore handling removed)
    if /i "!extension!"==".rar" (
        echo !baseName!| findstr /i /r /c:"^.*[.]part[0-9][0-9]*$" >nul
        if !ERRORLEVEL! EQU 0 (
            set "isMultipart=1"
            echo Detected multipart archive based on .partN suffix.

            echo !baseName!| findstr /i /r /c:"^.*[.]part0*1$" >nul
            if !ERRORLEVEL! EQU 0 (
                set "isFirstPart=1"
                echo It's the first part of the multipart archive.
                set "newBaseName=!baseName:.part1=!"
                set "newBaseName=!newBaseName:.part01=!"
                set "newBaseName=!newBaseName:.part001=!"
                echo New base name after removing ".partN": "!newBaseName!"
            ) else (
                echo It's NOT the first part of the multipart archive. Skipping this file.
                set "process=0"
            )
        ) else (
            echo Filename does NOT match any known multipart patterns.
        )
    )

    echo Final Base name: "!newBaseName!"
    echo Process flag: "!process!"

    :: Determine which tool supports this archive based on its extension
    set "Tool=none"
    if /i "!extension!"==".rar" set "Tool=unrar"
    if /i "!extension!"==".zip" set "Tool=7z"
    if /i "!extension!"==".7z" set "Tool=7z"
    if /i "!extension!"==".tar" set "Tool=7z"
    if /i "!extension!"==".gz" set "Tool=7z"
    if /i "!extension!"==".bz2" set "Tool=7z"
    if /i "!extension!"==".xz" set "Tool=7z"
    if /i "!extension!"==".tgz" set "Tool=7z"

    if "!Tool!"=="none" (
        echo File type !extension! is not supported yet. Skipping extraction.
        goto SkipExtraction
    )

    :: Continue only if this file needs to be processed
    if "!process!"=="1" (
        set "Success=0"

        :: Determine the output directory path
        if "!ExtractToSeparateFolders!"=="1" (
            set "OutputPath=Output\!newBaseName!\"
            echo Output path set to "!OutputPath!"
            if not exist "!OutputPath!" (
                echo Creating directory "!OutputPath!"
                mkdir "!OutputPath!"
                if !ERRORLEVEL! NEQ 0 (
                    echo Failed to create output directory "!OutputPath!". Skipping extraction.
                    set "process=0"
                )
            )
        ) else (
            set "OutputPath=Output\"
            echo Output path set to "!OutputPath!"
        )

        :: Skip extraction if directory creation failed
        if "!process!"=="0" (
            echo Skipping extraction due to previous errors.
            goto SkipExtraction
        )

        if /i "!Tool!"=="unrar" (
            :: Attempt extraction using unrar.exe
            if "!PasswordCount!"=="0" (
                echo No password provided, attempting to extract without a password...
                unrar.exe x "%%F" -o+ "!OutputPath!"
                if !ERRORLEVEL! EQU 0 (
                    echo Successfully extracted %%F
                    set "Success=1"
                    if "!DeleteAfterFinish!"=="1" (
                        echo Deleting %%F...
                        del "%%F"
                        if !ERRORLEVEL! EQU 0 (
                            echo Successfully deleted %%F.
                        ) else (
                            set /a DelFail+=1
                            echo Could not delete %%F.
                        )
                    )
                ) else (
                    echo Failed to extract %%F. A password may be required.
                )
            ) else (
                for /L %%P in (1,1,!PasswordCount!) do (
                    if "!Success!"=="0" (
                        set "password=!Password%%P!"
                        echo Trying password "!password!" for file "%%F"
                        unrar.exe x -p"!password!" "%%F" -o+ "!OutputPath!"
                        if !ERRORLEVEL! EQU 0 (
                            echo Successfully extracted %%F with password "!password!".
                            set "Success=1"
                            if "!DeleteAfterFinish!"=="1" (
                                echo Deleting %%F...
                                del "%%F"
                                if !ERRORLEVEL! EQU 0 (
                                    echo Successfully deleted %%F.
                                ) else (
                                    set /a DelFail+=1
                                    echo Could not delete %%F.
                                )
                            )
                        ) else (
                            echo Password "!password!" did not work for "%%F".
                        )
                    )
                )
            )

            if "!Success!"=="0" (
                echo All provided passwords for "%%F" are invalid or a password is required. Skipping this file.
                set /a DecompressFail+=1
            )
        ) else if /i "!Tool!"=="7z" (
            :: Attempt extraction using 7z.exe
            if "!PasswordCount!"=="0" (
                echo No password provided, attempting to extract without a password...
                7z.exe x "%%F" -o"!OutputPath!" -y
                if !ERRORLEVEL! EQU 0 (
                    echo Successfully extracted %%F
                    set "Success=1"
                    if "!DeleteAfterFinish!"=="1" (
                        echo Deleting %%F...
                        del "%%F"
                        if !ERRORLEVEL! EQU 0 (
                            echo Successfully deleted %%F.
                        ) else (
                            set /a DelFail+=1
                            echo Could not delete %%F.
                        )
                    )
                ) else (
                    echo Failed to extract %%F. A password may be required.
                )
            ) else (
                for /L %%P in (1,1,!PasswordCount!) do (
                    if "!Success!"=="0" (
                        set "password=!Password%%P!"
                        echo Trying password "!password!" for file "%%F"
                        7z.exe x -p"!password!" "%%F" -o"!OutputPath!" -y
                        if !ERRORLEVEL! EQU 0 (
                            echo Successfully extracted %%F with password "!password!".
                            set "Success=1"
                            if "!DeleteAfterFinish!"=="1" (
                                echo Deleting %%F...
                                del "%%F"
                                if !ERRORLEVEL! EQU 0 (
                                    echo Successfully deleted %%F.
                                ) else (
                                    set /a DelFail+=1
                                    echo Could not delete %%F.
                                )
                            )
                        ) else (
                            echo Password "!password!" did not work for "%%F".
                        )
                    )
                )
            )

            if "!Success!"=="0" (
                echo All provided passwords for "%%F" are invalid or a password is required. Skipping this file.
                set /a DecompressFail+=1
            )
        )
    )

    :SkipExtraction
    echo.
)

echo.
echo All files have been processed!
echo %DecompressFail% file(s) failed to extract.
echo %DelFail% file(s) could not be deleted.
echo.

echo Press any key to exit...
pause >nul
