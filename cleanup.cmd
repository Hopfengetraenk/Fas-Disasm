	@prompt -$G
	@cls
	@echo This will delete/clean all *.fct *.res and *.key files in:
	@echo %cd%
	@echo.
	@echo Press Ctrl+C to Cancel now. (Or anykey to contiune.)
	@pause
	@echo.
	@for /R %%i in  (*.f??.fct *.f??.res *.f??.key) do @call :DoDel "%%i"
	
	)
	::for /?
	@pause

@goto :EOF

:DoDel
	 @set File=%1
	 @echo %File:~-80%
	 @del %1 1>nul 
@goto :EOF