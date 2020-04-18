@echo off
c:\dos\rar.exe a c:\vb3\winready\winready.rar -s -m5 -ehs -va -v1457  -x@c:\vb3\winready\files.exc c:\vb3\winready\*.*
if not errorlevel 1 goto exito
echo Error en la compresión > c:\vb3\winready\error.err
goto end
:exito
echo Exito en la compresión >c:\vb3\winready\error.err
goto end
:end
