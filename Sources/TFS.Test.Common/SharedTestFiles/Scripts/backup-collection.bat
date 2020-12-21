echo off

set tc="C:\Program Files\Microsoft Team Foundation Server 12.0\Tools\TfsConfig.exe"
set backuproot=C:\Temp\BackupCollectionTest\BackupFolder
set dbserver=%COMPUTERNAME%
set collection=WordToTFS_TestCollection

set date=%date /t%
set time=%time /t%

for /F "usebackq tokens=1,2,3 delims=." %%I IN (`echo %%date%%`) do set date=%%K%%J%%I
for /F "usebackq tokens=1,2 delims=:" %%I IN (`echo %%time%%`) do set time=%%I%%J
set folder=%backuproot%\backups\%dbserver%\%date%-%time%
echo === backup to folder "%folder%" ===

set time=
set date=

mkdir "%folder%"

echo on

%tc% Collection /detach /CollectionName:"%collection%"

sqlcmd -S %dbserver% -i backup-collection.sql -v BackupFolder="%folder%" -v Database="Tfs_%collection%"

%tc% Collection /attach /CollectionName:"%collection%" /CollectionDB:"%dbServer%;Tfs_%collection%"

