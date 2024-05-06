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

	conda create --clone %ARCGIS_DEFAULT% --name %ENV_NAME% --pinned
	ECHO ^-^-^> %ENV_NAME% environment created
	)

call "C:\Program Files\ArcGIS\Pro\bin\Python\Scripts\activate.bat" %ENV_NAME% & (
	pip install shapely
	pip install geopandas==0.10.2
	pip install ezdxf
	pip install PyMuPdf --ignore-installed
	conda remove gdal --force --yes
	pip install GDAL-3.3.3-cp39-cp39-win_amd64.whl
	ECHO ^-^-^> shapely=2.0.1 geopandas=0.10.2 ezdxf gdal=3.3.3 installed
	)

CALL "%PRO_PATH%\Scripts\proswap.bat" %ENV_NAME%
ECHO ^-^-^> Pro env set

pause
exit