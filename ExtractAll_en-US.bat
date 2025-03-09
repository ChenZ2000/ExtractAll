@echo off
setlocal EnableDelayedExpansion

:: Initialize variables
set "i=1"
set "DecompressFail=0"
set "DelFail=0"
set "MultipartSuccessful="

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
for %%F in (*.rar *.zip *.7z *.tar *.gz *.bz2 *.xz *.tar.gz *.tgz *.zip.001 *.7z.001 *.z01 *.001) do (
    echo.
    echo ----------------------------------------
    echo Processing file: %%F
    set "process=1"
    set "baseName=%%~nF"
    set "extension=%%~xF"
    set "fullName=%%~nxF"
    set "fullPath=%%~fF"

    echo Initial Base name: "!baseName!"
    echo Extension: "!extension!"

    :: Initialize variables for multipart archive handling
    set "isMultipart=0"
    set "isFirstPart=0"
    set "newBaseName=!baseName!"
    set "multipartBaseName="
    set "multipartType="

    :: For RAR files, check for multipart archives with .partN pattern
    if /i "!extension!"==".rar" (
        echo !baseName!| findstr /i /r /c:"^.*[.]part[0-9][0-9]*$" >nul
        if !ERRORLEVEL! EQU 0 (
            set "isMultipart=1"
            set "multipartType=rar-part"
            echo Detected multipart archive based on .partN suffix.
            
            :: Extract the base name of the multipart archive (without .partN)
            for /f "tokens=1 delims=." %%a in ("!baseName!") do set "multipartBaseName=%%a"
            
            echo !baseName!| findstr /i /r /c:"^.*[.]part0*1$" >nul
            if !ERRORLEVEL! EQU 0 (
                set "isFirstPart=1"
                echo It's the first part of the multipart archive.
                set "newBaseName=!baseName:.part1=!"
                set "newBaseName=!newBaseName:.part01=!"
                set "newBaseName=!newBaseName:.part001=!"
                echo New base name after removing ".partN": "!newBaseName!"
            ) else (
                :: Check if we've already successfully extracted the first part of this multipart archive
                echo !MultipartSuccessful!|findstr /i "\[!multipartBaseName!-!multipartType!\]" >nul
                if !ERRORLEVEL! EQU 0 (
                    :: If we've successfully extracted the first part and need to delete, delete this part too
                    if "!DeleteAfterFinish!"=="1" (
                        echo Detected that this file is part of an already extracted multipart archive, preparing to delete...
                        echo Deleting %%F...
                        del "%%F"
                        if !ERRORLEVEL! EQU 0 (
                            echo Successfully deleted %%F.
                        ) else (
                            set /a DelFail+=1
                            echo Could not delete %%F.
                        )
                    )
                )
                echo It's NOT the first part of the multipart archive. Skipping extraction.
                set "process=0"
            )
        ) else (
            echo Filename does NOT match any known RAR multipart patterns.
        )
    )
    
    :: Check for .z01, .z02, etc. ZIP volume formats
    if /i "!extension!"==".z01" (
        set "isMultipart=1"
        set "multipartType=zip-part"
        echo Detected .z01 format ZIP volume.
        
        :: Try to find the corresponding .zip main file
        set "mainZipFile=!baseName!.zip"
        if exist "!mainZipFile!" (
            echo Found corresponding main ZIP file: !mainZipFile!
            :: Here z01 is not the first part, the main ZIP file is the first part
            set "multipartBaseName=!baseName!"
            
            :: Check if we've already successfully extracted this multipart archive series
            echo !MultipartSuccessful!|findstr /i "\[!multipartBaseName!-!multipartType!\]" >nul
            if !ERRORLEVEL! EQU 0 (
                :: If we've successfully extracted the main file and need to delete, delete the volume file too
                if "!DeleteAfterFinish!"=="1" (
                    echo Detected that this file is part of an already extracted ZIP volume, preparing to delete...
                    echo Deleting %%F...
                    del "%%F"
                    if !ERRORLEVEL! EQU 0 (
                        echo Successfully deleted %%F.
                    ) else (
                        set /a DelFail+=1
                        echo Could not delete %%F.
                    )
                )
            )
            :: Skip extraction processing anyway, as we only process the main file
            set "process=0"
        ) else (
            echo Could not find corresponding main ZIP file, cannot process this volume file.
            set "process=0"
        )
    )
    
    :: Check for .001, .002, etc. generic volume formats
    if /i "!extension!"==".001" (
        set "isMultipart=1"
        set "multipartType=generic-part"
        echo Detected .001 format volume.
        
        :: Extract base name from filename
        set "multipartBaseName=!baseName!"
        
        :: This is the first part
        set "isFirstPart=1"
        echo This is the first part of the multipart archive.
        
        :: Try to determine file type
        :: For simplicity, we'll assume it's a 7z or zip format
        :: In a real implementation, we should check file headers to determine type
        set "Tool=7z"
    )
    
    :: Check for .zip.001, .7z.001 formats
    if "!fullName:~-6!"==".001" (
        :: Check the extension before the .001, like .7z.001
        if "!fullName:~-8,2!"==".7" (
            set "isMultipart=1"
            set "multipartType=7z-part"
            echo Detected .7z.001 format 7Z volume.
            
            :: Extract base name from filename (remove .7z.001 part)
            set "multipartBaseName=!baseName:~0,-4!"
            
            :: This is the first part
            set "isFirstPart=1"
            echo This is the first part of the multipart archive.
        ) else if "!fullName:~-9,4!"==".zip" (
            set "isMultipart=1"
            set "multipartType=zip-part"
            echo Detected .zip.001 format ZIP volume.
            
            :: Extract base name from filename (remove .zip.001 part)
            set "multipartBaseName=!baseName:~0,-4!"
            
            :: This is the first part
            set "isFirstPart=1"
            echo This is the first part of the multipart archive.
        ) else (
            echo Cannot determine the volume type of !fullName!.
            set "isMultipart=1"
            set "multipartType=unknown-part"
            set "multipartBaseName=!baseName!"
            set "isFirstPart=1"
        )
    )
    
    :: Check for ZIP files, might be the main file of a multipart ZIP
    if /i "!extension!"==".zip" (
        :: Check if there's a corresponding .z01 volume file
        set "potentialPart=!baseName!.z01"
        if exist "!potentialPart!" (
            set "isMultipart=1"
            set "multipartType=zip-part"
            set "isFirstPart=1"
            set "multipartBaseName=!baseName!"
            echo Detected that this is the main file of a ZIP multipart archive, corresponding volume files exist.
        )
    )
    
    :: Check for 7z files, might be the main file of a multipart 7z (though 7z typically uses .001 format rather than specialized extensions)
    if /i "!extension!"==".7z" (
        :: Check if there's a corresponding .001 volume file
        set "potentialPart=!baseName!.001"
        if exist "!potentialPart!" (
            set "isMultipart=1"
            set "multipartType=7z-part"
            set "isFirstPart=1"
            set "multipartBaseName=!baseName!"
            echo Detected that this is the main file of a 7Z multipart archive, corresponding volume files exist.
        )
    )

    echo Multipart status: !isMultipart! Is first part: !isFirstPart! Base name: "!multipartBaseName!"
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
    if /i "!extension!"==".001" set "Tool=7z"
    if /i "!extension!"==".z01" set "Tool=skip"
    
    :: Handle .zip.001 and .7z.001 formats
    if "!fullName:~-6!"==".001" set "Tool=7z"

    if "!Tool!"=="none" (
        echo File type !extension! is not supported yet. Skipping extraction.
        goto SkipExtraction
    )
    
    if "!Tool!"=="skip" (
        echo This is a volume file that doesn't need separate processing. Skipping extraction.
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
                    :: If this is the first part of a multipart archive, record the successful extraction
                    if "!isMultipart!"=="1" if "!isFirstPart!"=="1" (
                        set "MultipartSuccessful=!MultipartSuccessful![!multipartBaseName!-!multipartType!]"
                        echo Recorded multipart archive [!multipartBaseName!-!multipartType!] as successfully extracted
                    )
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
                            :: If this is the first part of a multipart archive, record the successful extraction
                            if "!isMultipart!"=="1" if "!isFirstPart!"=="1" (
                                set "MultipartSuccessful=!MultipartSuccessful![!multipartBaseName!-!multipartType!]"
                                echo Recorded multipart archive [!multipartBaseName!-!multipartType!] as successfully extracted
                            )
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
                    :: If this is the first part of a multipart archive, record the successful extraction
                    if "!isMultipart!"=="1" if "!isFirstPart!"=="1" (
                        set "MultipartSuccessful=!MultipartSuccessful![!multipartBaseName!-!multipartType!]"
                        echo Recorded multipart archive [!multipartBaseName!-!multipartType!] as successfully extracted
                    )
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
                            :: If this is the first part of a multipart archive, record the successful extraction
                            if "!isMultipart!"=="1" if "!isFirstPart!"=="1" (
                                set "MultipartSuccessful=!MultipartSuccessful![!multipartBaseName!-!multipartType!]"
                                echo Recorded multipart archive [!multipartBaseName!-!multipartType!] as successfully extracted
                            )
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

:: Process again to delete any remaining volume files (for special cases)
if "!DeleteAfterFinish!"=="1" (
    echo.
    echo Checking for any unprocessed volume files that need to be deleted...
    
    for %%F in (*.z?? *.????.??? *.???.??? *.??.??? *.?.??? *.???.?? *.??.?? *.?.??) do (
        set "fullName=%%~nxF"
        set "baseName=%%~nF"
        set "extension=%%~xF"
        set "multipartBaseName="
        set "multipartType="
        set "shouldDelete=0"
        
        :: Check for .z01, .z02 formats
        if "!extension:~0,2!"==".z" (
            set "multipartBaseName=!baseName!"
            set "multipartType=zip-part"
            echo !MultipartSuccessful!|findstr /i "\[!multipartBaseName!-!multipartType!\]" >nul
            if !ERRORLEVEL! EQU 0 set "shouldDelete=1"
        )
        
        :: Check for .001, .002 formats
        if "!extension:~0,1!"=="." (
            if "!extension:~1,1!" GEQ "0" if "!extension:~1,1!" LEQ "9" (
                :: Likely a numeric extension
                set "multipartBaseName=!baseName!"
                
                :: Try different multipart types
                for %%t in (generic-part 7z-part zip-part) do (
                    set "multipartType=%%t"
                    echo !MultipartSuccessful!|findstr /i "\[!multipartBaseName!-!multipartType!\]" >nul
                    if !ERRORLEVEL! EQU 0 set "shouldDelete=1"
                )
            )
        )
        
        :: Check for .part2.rar, .part3.rar formats
        echo !fullName!| findstr /i /r /c:"^.*[.]part[0-9][0-9]*[.]rar$" >nul
        if !ERRORLEVEL! EQU 0 (
            for /f "tokens=1 delims=." %%a in ("!baseName!") do set "multipartBaseName=%%a"
            set "multipartType=rar-part"
            echo !MultipartSuccessful!|findstr /i "\[!multipartBaseName!-!multipartType!\]" >nul
            if !ERRORLEVEL! EQU 0 set "shouldDelete=1"
        )
        
        if "!shouldDelete!"=="1" (
            echo Detected volume file "%%F" belongs to a successfully extracted series, preparing to delete...
            del "%%F"
            if !ERRORLEVEL! EQU 0 (
                echo Successfully deleted %%F.
            ) else (
                set /a DelFail+=1
                echo Could not delete %%F.
            )
        )
    )
)

echo.
echo All files have been processed!
echo %DecompressFail% file(s) failed to extract.
echo %DelFail% file(s) could not be deleted.
echo.

echo Press any key to exit...
pause >nul