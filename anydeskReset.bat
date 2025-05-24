@echo off
REM Caminha até a pasta desejada
cd "C:\ProgramData\AnyDesk"
echo programdata
echo %USERPROFILE%
REM Verifica e apaga o arquivo services.conf, se existir
if exist "service.conf" (
    del "service.conf"
    echo service.conf apagado com sucesso.
) else (
    echo service.conf nao encontrado.
)

REM Verifica e apaga o arquivo system.conf, se existir
if exist "system.conf" (
    del "system.conf"
    echo system.conf apagado com sucesso.
) else (
    echo system.conf nao encontrado.
)

cd "%USERPROFILE%\AppData\Roaming\AnyDesk"
echo ""
echo ""
echo appdata
echo ""
REM Exclui cada arquivo se existir
if exist "service.conf" (
    del "service.conf"
    echo service.conf excluido com sucesso.
) else (
    echo service.conf nao encontrado.
)

if exist "system.conf" (
    del "system.conf"
    echo system.conf excluido com sucesso.
) else (
    echo system.conf nao encontrado.
)

REM Verifica e apaga o diretório Thumbnails se existir
if exist "Thumbnails" (
    rmdir /s /q "Thumbnails"
    echo Thumbnails excluido com sucesso.
) else (
    echo Thumbnails nao encontrado.
)

if exist "ad.trace" (
    del "ad.trace"
    echo ad.trace excluido com sucesso.
) else (
    echo ad.trace nao encontrado.
)

if exist "user.conf" (
    del "user.conf"
    echo user.conf excluido com sucesso.
) else (
    echo user.conf nao encontrado.
)

echo Operacao concluida.
pause
