cscript stop.vbs
del %systemroot%\System32\spool\printers\* /Q /F /S
cscript start.vbs
