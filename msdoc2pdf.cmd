@echo off
rem Este script converte arquivos MSWord do tipo .doc que estiverem pasta 
rem   informada pela variavel !dir_in! e os converte para PDF colocando-os 
rem   na pasta informada pela variavel !dir_out!
rem   Por padrao !dir_out! deveria estar dentro de !dir_in! com o nome 
rem   sugestivo de "pdf" para simplificar operações de arquivos.
rem Os pdfs gerados estarao protegidos contra impressao
rem Dependencias necessárias:
rem 1) Requer o MSOFFice 2016+ instalado e registrado, versão com licença
rem   expirada nao funcionará
rem 2) Requer o programa PDFTK Server instalado, este programa é gratuito.
rem   Todo: substituir no futro o PDFTK por QPDF que é opensource;
rem Observacoes: O MSOffice deve estar instalado e ativado e também deve ser
rem   o aplicativo padrão para doc e docx. Os arquivos a serem convertidos
rem   nao podem ter espaço no meio do nome.
rem Autor: Gladiston Santana <gladiston.santana [em] gmail.com>
rem Data: 02/01/2016
rem
setlocal enableextensions enabledelayedexpansion
set pdfpass=qualsenhausar
set dir_in=C:\Downloads
set dir_out=%dir_in%\pdf
set arq_lista=%dir_in%\lista-pdf.txt
set dir_app=C:\Program Files (x86)\PDFtk Server\bin
set path=%path%;%dir_app%;
set EXIT_CODE=0
rem ========================================
rem nao faça alteracoes deste ponto em diante
rem ========================================
if not exist "%dir_out%" mkdir -p "%dir_out%" 
if not exist "%dir_out%" goto dir_out_error
rem if not exist "%exe_docto%" goto exe_docto_not_exist
if not exist "%dir_app%\pdftk.exe" goto exe_app_not_exist
rem --bookmarksource --pdf-BookmarkSource     
rem     PDF conversions can take their bookmarks from
rem    WordBookmarks, WordHeadings (default) or None
set pdf_bookmark=--pdf-BookmarkSource WordHeadings 
echo =================================================
echo Convertendo arquivos do MSWord para PDF em lote
echo Atenção: Arquivos nao podem ter espaços no nome.
echo Doc: !dir_in!
echo PDF: !dir_out!
echo =================================================

"%exe_docto%" -f "%dir_in%\" -O "%dir_out%\" -T wdFormatPDF  -OX .pdf %pdf_bookmark% >%arq_lista%
set EXIT_CODE=%ERRORLEVEL%
IF %EXIT_CODE% NEQ 0 goto fim
rem colocando senha nos PDFs gerados
for /f "tokens=*" %%A in (!arq_lista!) do (
  set file_in=%%A
  for /f "tokens=2" %%i in ("!file_in!") do set file_in=%%i
  rem removendo os espaços
  set file_in=!file_in: =!
  if exist "!file_in!" (
    set file_out=!file_in!
    set file_tmp=!file_in!
    rem removendo a extensao
    for /f "tokens=1 delims=." %%i in ("!file_tmp!") do set file_tmp=%%i
    set file_tmp=!file_tmp!-not-encripted.pdf

    rem renomeando a origem para o temporario
    move /y "!file_in!" "!file_tmp!" >nul

    if exist "!file_tmp!" (
      echo   De: "!file_tmp!" 
      echo Para: "!file_out!"
      pdftk ^
        "!file_tmp!" ^
        output "!file_out!" ^
		encrypt_40bit ^
        owner_pw "%pdfpass%" ^
        allow ScreenReaders ^
        allow ModifyAnnotations ^
		verbose		
      set EXIT_CODE=!ERRORLEVEL!
      if !EXIT_CODE! NEQ 0 (
        echo Falha na conversao error=!EXIT_CODE!     
        goto fim
      )
      rem apagando o temporario
      del /q /f "!file_tmp!"     
    )    
  )
)
set EXIT_CODE=!ERRORLEVEL!
IF !EXIT_CODE! NEQ 0 goto fim

rem apagando arquivo que continha lista
if exist "!arq_lista!" del /q /f "!arq_lista!" 

echo Ok, tudo pronto.
goto fim

:exe_app_not_exist
  echo Conversor PDFTK nao existe: "%dir_app%\pdftk.exe"
  pause
  goto fim

:dir_out_error
  echo Pasta nao existe: %dir_out%
  pause
  goto fim
  
:fim
