@echo off
setlocal EnableDelayedExpansion

:: 初始化变量
set "i=1"
set "DecompressFail=0"
set "DelFail=0"
set "MultipartSuccessful="

echo 版权信息 ↓↓↓
echo UnRar.exe © win.rar GmbH 保留所有权利。 ^<https://www.win-rar.com/^>
echo 7z.exe © Igor Pavlov. ^<https://www.7-zip.org/^>
echo 脚本由ChenZ编写并打包 ^<https://chenz.cloud^>
echo.
echo 欢迎使用ExtractAll^^!
:: 提示用户输入密码
:PasswordLoop
set /p Password%i%="请输入密码 %i% (留空进入下一步): "
if "!Password%i%!"=="" (
    set /a PasswordCount=!i!-1
    goto AskExtractOption
) else (
    set /a i+=1
    goto PasswordLoop
)

:AskExtractOption
echo.
:: 询问是否将每个压缩包解压到单独的子文件夹
echo 是否希望将每个压缩包解压到各自的子文件夹中？
set /p SeparateFolders=请输入Y或N并按回车: 
if /i "!SeparateFolders!"=="Y" (
    set "ExtractToSeparateFolders=1"
) else (
    set "ExtractToSeparateFolders=0"
)
echo.

:: 询问是否在解压后删除源压缩包
echo 是否希望在解压后删除压缩包？
set /p DeleteSource=请输入Y或N并按回车: 
if /i "!DeleteSource!"=="Y" (
    set "DeleteAfterFinish=1"
) else (
    set "DeleteAfterFinish=0"
)

echo.
echo 开始解压...
echo.

:: 创建输出目录（如果尚未存在）
if not exist "Output\" (
    mkdir "Output\"
    if !ERRORLEVEL! NEQ 0 (
        echo 创建输出目录 "Output\" 失败。正在退出。
        exit /b 1
    )
)

:: 处理当前目录下的所有压缩文件
for %%F in (*.rar *.zip *.7z *.tar *.gz *.bz2 *.xz *.tar.gz *.tgz *.zip.001 *.7z.001 *.z01 *.001) do (
    echo.
    echo ----------------------------------------
    echo 正在处理文件: %%F
    set "process=1"
    set "baseName=%%~nF"
    set "extension=%%~xF"
    set "fullName=%%~nxF"
    set "fullPath=%%~fF"

    echo 初始基本名称: "!baseName!"
    echo 扩展名: "!extension!"

    :: 初始化多卷压缩包处理变量
    set "isMultipart=0"
    set "isFirstPart=0"
    set "newBaseName=!baseName!"
    set "multipartBaseName="
    set "multipartType="

    :: 对于RAR文件，检查带有.partN模式的多卷压缩包
    if /i "!extension!"==".rar" (
        echo !baseName!| findstr /i /r /c:"^.*[.]part[0-9][0-9]*$" >nul
        if !ERRORLEVEL! EQU 0 (
            set "isMultipart=1"
            set "multipartType=rar-part"
            echo 基于.partN后缀检测到多卷RAR压缩包。
            
            :: 提取多卷压缩包的基本名称（不含.partN部分）
            for /f "tokens=1 delims=." %%a in ("!baseName!") do set "multipartBaseName=%%a"
            
            echo !baseName!| findstr /i /r /c:"^.*[.]part0*1$" >nul
            if !ERRORLEVEL! EQU 0 (
                set "isFirstPart=1"
                echo 这是多卷压缩包的第一部分。
                set "newBaseName=!baseName:.part1=!"
                set "newBaseName=!newBaseName:.part01=!"
                set "newBaseName=!newBaseName:.part001=!"
                echo 移除".partN"后的新基本名称: "!newBaseName!"
            ) else (
                :: 检查是否已经成功解压了这个多卷压缩包系列的第一部分
                echo !MultipartSuccessful!|findstr /i "\[!multipartBaseName!-!multipartType!\]" >nul
                if !ERRORLEVEL! EQU 0 (
                    :: 如果我们成功解压了第一部分且需要删除，那么也删除后续部分
                    if "!DeleteAfterFinish!"=="1" (
                        echo 检测到此文件是已成功解压的多卷压缩包的一部分，准备删除...
                        echo 正在删除 %%F...
                        del "%%F"
                        if !ERRORLEVEL! EQU 0 (
                            echo 成功删除 %%F。
                        ) else (
                            set /a DelFail+=1
                            echo 无法删除 %%F。
                        )
                    )
                )
                echo 这不是多卷压缩包的第一部分。跳过解压处理。
                set "process=0"
            )
        ) else (
            echo 文件名不匹配任何已知的RAR多卷模式。
        )
    )
    
    :: 检查.z01, .z02等zip分卷模式
    if /i "!extension!"==".z01" (
        set "isMultipart=1"
        set "multipartType=zip-part"
        echo 检测到.z01格式的ZIP分卷。
        
        :: 尝试查找对应的.zip主文件
        set "mainZipFile=!baseName!.zip"
        if exist "!mainZipFile!" (
            echo 找到对应的主ZIP文件: !mainZipFile!
            :: 这里z01不是第一部分，而是主ZIP文件是第一部分
            set "multipartBaseName=!baseName!"
            
            :: 检查是否已经成功解压了这个多卷压缩包系列
            echo !MultipartSuccessful!|findstr /i "\[!multipartBaseName!-!multipartType!\]" >nul
            if !ERRORLEVEL! EQU 0 (
                :: 如果我们成功解压了主文件且需要删除，那么也删除分卷文件
                if "!DeleteAfterFinish!"=="1" (
                    echo 检测到此文件是已成功解压的ZIP分卷的一部分，准备删除...
                    echo 正在删除 %%F...
                    del "%%F"
                    if !ERRORLEVEL! EQU 0 (
                        echo 成功删除 %%F。
                    ) else (
                        set /a DelFail+=1
                        echo 无法删除 %%F。
                    )
                )
            )
            :: 无论如何都跳过解压处理，因为我们只处理主文件
            set "process=0"
        ) else (
            echo 找不到对应的主ZIP文件，无法处理这个分卷文件。
            set "process=0"
        )
    )
    
    :: 检查.001, .002等通用分卷模式
    if /i "!extension!"==".001" (
        set "isMultipart=1"
        set "multipartType=generic-part"
        echo 检测到.001格式的分卷。
        
        :: 从文件名中提取基本名称
        set "multipartBaseName=!baseName!"
        
        :: 这是第一部分
        set "isFirstPart=1"
        echo 这是多卷压缩包的第一部分。
        
        :: 尝试确定文件类型
        :: 复制文件头几个字节到临时文件中分析
        :: 简化处理，假设它是7z或zip格式
        :: 实际上这里应该检查文件头来确定类型，但为简化起见我们直接处理
        set "Tool=7z"
    )
    
    :: 检查.zip.001, .7z.001等格式
    if "!fullName:~-6!"==".001" (
        :: 检查扩展名前两位，如.7z.001
        if "!fullName:~-8,2!"==".7" (
            set "isMultipart=1"
            set "multipartType=7z-part"
            echo 检测到.7z.001格式的7Z分卷。
            
            :: 从文件名中提取基本名称（去掉.7z.001部分）
            set "multipartBaseName=!baseName:~0,-4!"
            
            :: 这是第一部分
            set "isFirstPart=1"
            echo 这是多卷压缩包的第一部分。
        ) else if "!fullName:~-9,4!"==".zip" (
            set "isMultipart=1"
            set "multipartType=zip-part"
            echo 检测到.zip.001格式的ZIP分卷。
            
            :: 从文件名中提取基本名称（去掉.zip.001部分）
            set "multipartBaseName=!baseName:~0,-4!"
            
            :: 这是第一部分
            set "isFirstPart=1"
            echo 这是多卷压缩包的第一部分。
        ) else (
            echo 无法确定 !fullName! 的分卷类型。
            set "isMultipart=1"
            set "multipartType=unknown-part"
            set "multipartBaseName=!baseName!"
            set "isFirstPart=1"
        )
    )
    
    :: 检查zip文件，可能是多卷zip的主文件
    if /i "!extension!"==".zip" (
        :: 检查是否有对应的.z01分卷文件
        set "potentialPart=!baseName!.z01"
        if exist "!potentialPart!" (
            set "isMultipart=1"
            set "multipartType=zip-part"
            set "isFirstPart=1"
            set "multipartBaseName=!baseName!"
            echo 检测到这是ZIP多卷压缩包的主文件，存在对应的分卷文件。
        )
    )
    
    :: 检查7z文件，可能是多卷7z的主文件（虽然7z通常使用.001格式而不是专用扩展名）
    if /i "!extension!"==".7z" (
        :: 检查是否有对应的.001分卷文件
        set "potentialPart=!baseName!.001"
        if exist "!potentialPart!" (
            set "isMultipart=1"
            set "multipartType=7z-part"
            set "isFirstPart=1"
            set "multipartBaseName=!baseName!"
            echo 检测到这是7Z多卷压缩包的主文件，存在对应的分卷文件。
        )
    )

    echo 多卷状态: !isMultipart! 是首卷: !isFirstPart! 基本名称: "!multipartBaseName!"
    echo 最终基本名称: "!newBaseName!"
    echo 处理标志: "!process!"

    :: 根据扩展名确定支持此压缩包的工具
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
    
    :: 处理.zip.001和.7z.001格式
    if "!fullName:~-6!"==".001" set "Tool=7z"

    if "!Tool!"=="none" (
        echo 文件类型 !extension! 尚不支持。跳过解压。
        goto SkipExtraction
    )
    
    if "!Tool!"=="skip" (
        echo 这是一个分卷文件，不需要单独处理。跳过解压。
        goto SkipExtraction
    )

    :: 仅在需要处理此文件时继续
    if "!process!"=="1" (
        set "Success=0"

        :: 确定输出目录路径
        if "!ExtractToSeparateFolders!"=="1" (
            set "OutputPath=Output\!newBaseName!\"
            echo 输出路径设置为 "!OutputPath!"
            if not exist "!OutputPath!" (
                echo 正在创建目录 "!OutputPath!"
                mkdir "!OutputPath!"
                if !ERRORLEVEL! NEQ 0 (
                    echo 创建输出目录 "!OutputPath!" 失败。跳过解压。
                    set "process=0"
                )
            )
        ) else (
            set "OutputPath=Output\"
            echo 输出路径设置为 "!OutputPath!"
        )

        :: 如果目录创建失败，跳过解压
        if "!process!"=="0" (
            echo 由于之前的错误，跳过解压。
            goto SkipExtraction
        )

        if /i "!Tool!"=="unrar" (
            :: 尝试使用unrar.exe解压
            if "!PasswordCount!"=="0" (
                echo 未提供密码，尝试无密码解压...
                unrar.exe x "%%F" -o+ "!OutputPath!"
                if !ERRORLEVEL! EQU 0 (
                    echo 成功解压 %%F
                    set "Success=1"
                    :: 如果是多卷压缩包的第一部分，记录成功解压的基本名称
                    if "!isMultipart!"=="1" if "!isFirstPart!"=="1" (
                        set "MultipartSuccessful=!MultipartSuccessful![!multipartBaseName!-!multipartType!]"
                        echo 记录多卷压缩包 [!multipartBaseName!-!multipartType!] 已成功解压
                    )
                    if "!DeleteAfterFinish!"=="1" (
                        echo 正在删除 %%F...
                        del "%%F"
                        if !ERRORLEVEL! EQU 0 (
                            echo 成功删除 %%F。
                        ) else (
                            set /a DelFail+=1
                            echo 无法删除 %%F。
                        )
                    )
                ) else (
                    echo 解压 %%F 失败。可能需要密码。
                )
            ) else (
                for /L %%P in (1,1,!PasswordCount!) do (
                    if "!Success!"=="0" (
                        set "password=!Password%%P!"
                        echo 正在尝试密码 "!password!" 解压文件 "%%F"
                        unrar.exe x -p"!password!" "%%F" -o+ "!OutputPath!"
                        if !ERRORLEVEL! EQU 0 (
                            echo 使用密码 "!password!" 成功解压 %%F。
                            set "Success=1"
                            :: 如果是多卷压缩包的第一部分，记录成功解压的基本名称
                            if "!isMultipart!"=="1" if "!isFirstPart!"=="1" (
                                set "MultipartSuccessful=!MultipartSuccessful![!multipartBaseName!-!multipartType!]"
                                echo 记录多卷压缩包 [!multipartBaseName!-!multipartType!] 已成功解压
                            )
                            if "!DeleteAfterFinish!"=="1" (
                                echo 正在删除 %%F...
                                del "%%F"
                                if !ERRORLEVEL! EQU 0 (
                                    echo 成功删除 %%F。
                                ) else (
                                    set /a DelFail+=1
                                    echo 无法删除 %%F。
                                )
                            )
                        ) else (
                            echo 密码 "!password!" 不适用于 "%%F"。
                        )
                    )
                )
            )

            if "!Success!"=="0" (
                echo 提供的所有密码对 "%%F" 无效或需要密码。跳过此文件。
                set /a DecompressFail+=1
            )
        ) else if /i "!Tool!"=="7z" (
            :: 尝试使用7z.exe解压
            if "!PasswordCount!"=="0" (
                echo 未提供密码，尝试无密码解压...
                7z.exe x "%%F" -o"!OutputPath!" -y
                if !ERRORLEVEL! EQU 0 (
                    echo 成功解压 %%F
                    set "Success=1"
                    :: 如果是多卷压缩包的第一部分，记录成功解压的基本名称
                    if "!isMultipart!"=="1" if "!isFirstPart!"=="1" (
                        set "MultipartSuccessful=!MultipartSuccessful![!multipartBaseName!-!multipartType!]"
                        echo 记录多卷压缩包 [!multipartBaseName!-!multipartType!] 已成功解压
                    )
                    if "!DeleteAfterFinish!"=="1" (
                        echo 正在删除 %%F...
                        del "%%F"
                        if !ERRORLEVEL! EQU 0 (
                            echo 成功删除 %%F。
                        ) else (
                            set /a DelFail+=1
                            echo 无法删除 %%F。
                        )
                    )
                ) else (
                    echo 解压 %%F 失败。可能需要密码。
                )
            ) else (
                for /L %%P in (1,1,!PasswordCount!) do (
                    if "!Success!"=="0" (
                        set "password=!Password%%P!"
                        echo 正在尝试密码 "!password!" 解压文件 "%%F"
                        7z.exe x -p"!password!" "%%F" -o"!OutputPath!" -y
                        if !ERRORLEVEL! EQU 0 (
                            echo 使用密码 "!password!" 成功解压 %%F。
                            set "Success=1"
                            :: 如果是多卷压缩包的第一部分，记录成功解压的基本名称
                            if "!isMultipart!"=="1" if "!isFirstPart!"=="1" (
                                set "MultipartSuccessful=!MultipartSuccessful![!multipartBaseName!-!multipartType!]"
                                echo 记录多卷压缩包 [!multipartBaseName!-!multipartType!] 已成功解压
                            )
                            if "!DeleteAfterFinish!"=="1" (
                                echo 正在删除 %%F...
                                del "%%F"
                                if !ERRORLEVEL! EQU 0 (
                                    echo 成功删除 %%F。
                                ) else (
                                    set /a DelFail+=1
                                    echo 无法删除 %%F。
                                )
                            )
                        ) else (
                            echo 密码 "!password!" 不适用于 "%%F"。
                        )
                    )
                )
            )

            if "!Success!"=="0" (
                echo 提供的所有密码对 "%%F" 无效或需要密码。跳过此文件。
                set /a DecompressFail+=1
            )
        )
    )

    :SkipExtraction
    echo.
)

:: 再次处理以删除剩余的分卷文件（对于某些特殊情况）
if "!DeleteAfterFinish!"=="1" (
    echo.
    echo 检查是否有未处理的分卷文件需要删除...
    
    for %%F in (*.z?? *.????.??? *.???.??? *.??.??? *.?.??? *.???.?? *.??.?? *.?.??) do (
        set "fullName=%%~nxF"
        set "baseName=%%~nF"
        set "extension=%%~xF"
        set "multipartBaseName="
        set "multipartType="
        set "shouldDelete=0"
        
        :: 检查.z01, .z02等格式
        if "!extension:~0,2!"==".z" (
            set "multipartBaseName=!baseName!"
            set "multipartType=zip-part"
            echo !MultipartSuccessful!|findstr /i "\[!multipartBaseName!-!multipartType!\]" >nul
            if !ERRORLEVEL! EQU 0 set "shouldDelete=1"
        )
        
        :: 检查.001, .002等格式
        if "!extension:~0,1!"=="." (
            if "!extension:~1,1!" GEQ "0" if "!extension:~1,1!" LEQ "9" (
                :: 可能是数字扩展名
                set "multipartBaseName=!baseName!"
                
                :: 尝试不同的多卷类型
                for %%t in (generic-part 7z-part zip-part) do (
                    set "multipartType=%%t"
                    echo !MultipartSuccessful!|findstr /i "\[!multipartBaseName!-!multipartType!\]" >nul
                    if !ERRORLEVEL! EQU 0 set "shouldDelete=1"
                )
            )
        )
        
        :: 检查.part2.rar, .part3.rar等格式
        echo !fullName!| findstr /i /r /c:"^.*[.]part[0-9][0-9]*[.]rar$" >nul
        if !ERRORLEVEL! EQU 0 (
            for /f "tokens=1 delims=." %%a in ("!baseName!") do set "multipartBaseName=%%a"
            set "multipartType=rar-part"
            echo !MultipartSuccessful!|findstr /i "\[!multipartBaseName!-!multipartType!\]" >nul
            if !ERRORLEVEL! EQU 0 set "shouldDelete=1"
        )
        
        if "!shouldDelete!"=="1" (
            echo 检测到分卷文件 "%%F" 属于已成功解压的系列，准备删除...
            del "%%F"
            if !ERRORLEVEL! EQU 0 (
                echo 成功删除 %%F。
            ) else (
                set /a DelFail+=1
                echo 无法删除 %%F。
            )
        )
    )
)

echo.
echo 所有文件处理完毕！
echo %DecompressFail% 个文件解压失败。
echo %DelFail% 个文件无法删除。
echo.

echo 按任意键退出...
pause >nul
