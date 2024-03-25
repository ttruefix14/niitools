:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: VARIABLES                                                                    :
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
@echo off

SETLOCAL
SET "ENV_NAME=arcgispro-py3-clone"
SET "ARCGIS_DEFAULT=arcgispro-py3"

SET "PRO_PATH=C:\Program Files\ArcGIS\Pro\bin\Python"
SET "CLONED_PATH=%LOCALAPPDATA%\ESRI\conda\envs"
REM Set the default script path for this iteration. Maybe fixes jupyter notebook install errors?
SET PATH=%PATH%;"C:\Program Files\ArcGIS\Pro\bin\Python\Scripts"

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: COMMANDS                                                                     :
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

REM start by activating the arcgis pro conda env
call "C:\Program Files\ArcGIS\Pro\bin\Python\Scripts\activate.bat" & (

	activate arcgispro-py3
	conda list
	)
pause
exit