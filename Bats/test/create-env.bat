:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: VARIABLES                                                                    :
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
@echo off

SETLOCAL
SET "ENV_NAME=assessors-env"
SET "ARCGIS_DEFAULT=arcgispro-py3"
SET "ARCGIS_CLONE=arcgispro-py3-clone"

SET "PRO_PATH=C:\Program Files\ArcGIS\Pro\bin\Python"
SET "CLONED_PATH=%LOCALAPPDATA%\ESRI\conda\envs"
REM Set the default script path for this iteration. Maybe fixes jupyter notebook install errors?
SET PATH=%PATH%;"C:\Program Files\ArcGIS\Pro\bin\Python\Scripts"

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: COMMANDS                                                                     :
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

REM start by activating the arcgis pro conda env
call "C:\Program Files\ArcGIS\Pro\bin\Python\Scripts\activate.bat" & (
	REM export the cloned arcgispro-py3-clone environment that contains the additional packages installed outside of arcgis pro
		conda env export -n %ARCGIS_CLONE% > "%CLONED_PATH%\%ARCGIS_CLONE%\environment.yml"	
		ECHO ^-^-^> %ARCGIS_CLONE% conda environment exported to "%CLONED_PATH%\%ARCGIS_CLONE%\environment.yml"
		
	REM clone the deafult arcgispro-py3 environment to the new environment name
		REM CALL conda config --add channels conda-forge
		REM CALL conda config --add channels esri
		conda create --clone %ARCGIS_DEFAULT% --name %ENV_NAME% --pinned
		ECHO ^-^-^> %ENV_NAME% environment created
		
	REM update the new environment with packages that are not installed from the environment.yml created earlier.
	    REM conda env list
		REM conda env update -n %ENV_NAME% -f "%PRO_PATH%\envs\%ARCGIS_DEFAULT%\environment.yml"
		conda env update -n %ENV_NAME% -f "%CLONED_PATH%\%ARCGIS_CLONE%\environment.yml"
		ECHO ^-^-^> %ENV_NAME% environment updated
	
	REM set the active environment in Pro to the new environment
		CALL "%PRO_PATH%\Scripts\proswap.bat" %ENV_NAME%
		ECHO ^-^-^> Pro env set
	)
pause
exit