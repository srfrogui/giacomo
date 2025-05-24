@echo off
setlocal

:: Verifica se foi passado argumento, senão pede com janelinha
if "%~1"=="" (
    for /f "tokens=*" %%a in ('mshta "javascript:var commit=prompt('Digite a mensagem do commit:', '');new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(commit);close()"') do set "COMMIT_MSG=%%a"
) else (
    set "COMMIT_MSG=%~1"
)

:: Verifica se a mensagem foi definida
if "%COMMIT_MSG%"=="" (
    echo Nenhuma mensagem de commit foi fornecida.
    exit /b 1
)

:: Adiciona arquivos
git add .

:: Faz commit com a mensagem
git commit -m "%COMMIT_MSG%"

:: Faz push para a branch main
git push origin main

:: Só pausa se SKIP_PAUSE não estiver definido
if not defined SKIP_PAUSE pause
