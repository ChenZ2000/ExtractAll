@echo off
setlocal EnableDelayedExpansion

:: 初始化变量
set "i=1"
set "DecompressFail=0"
set "DelFail=0"

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
for %%F in (*.rar *.zip *.7z *.tar *.gz *.bz2 *.xz *.tar.gz *.tgz) do (
    echo.
    echo ----------------------------------------
    echo 正在处理文件: %%F
    set "process=1"
    set "baseName=%%~nF"
    set "extension=%%~xF"

    echo 初始基本名称: "!baseName!"
    echo 扩展名: "!extension!"

    :: 初始化多卷压缩包处理变量
    set "isMultipart=0"
    set "isFirstPart=0"
    set "newBaseName=!baseName!"

    :: 对于RAR文件，检查带有.partN模式的多卷压缩包（已移除下划线处理）
    if /i "!extension!"==".rar" (
        echo !baseName!| findstr /i /r /c:"^.*[.]part[0-9][0-9]*$" >nul
        if !ERRORLEVEL! EQU 0 (
            set "isMultipart=1"
            echo 基于.partN后缀检测到多卷压缩包。

            echo !baseName!| findstr /i /r /c:"^.*[.]part0*1$" >nul
            if !ERRORLEVEL! EQU 0 (
                set "isFirstPart=1"
                echo 这是多卷压缩包的第一部分。
                set "newBaseName=!baseName:.part1=!"
                set "newBaseName=!newBaseName:.part01=!"
                set "newBaseName=!newBaseName:.part001=!"
                echo 移除".partN"后的新基本名称: "!newBaseName!"
            ) else (
                echo 这不是多卷压缩包的第一部分。跳过此文件。
                set "process=0"
            )
        ) else (
            echo 文件名不匹配任何已知的多卷模式。
        )
    )

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

    if "!Tool!"=="none" (
        echo 文件类型 !extension! 尚不支持。跳过解压。
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

echo.
echo 所有文件处理完毕！
echo %DecompressFail% 个文件解压失败。
echo %DelFail% 个文件无法删除。
echo.

echo 按任意键退出...
pause >nul