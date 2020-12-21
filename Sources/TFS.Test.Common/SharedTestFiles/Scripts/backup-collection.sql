DECLARE @name VARCHAR(256)     -- database name  
DECLARE @path VARCHAR(256)     -- path for backup files  
DECLARE @fileName VARCHAR(256) -- filename for backup  
 
-- specify database backup directory - $(BackupFolder) provided by sqlcmd -a BackupFolder="<...>"
SET @path = '$(BackupFolder)' 
 
-- specify database name - $(Database) provided by sqlcmd -a Database="<...>"
SET @name = '$(Database)' 

SET @fileName = @path + '\' + @name + '.bak'  
BACKUP DATABASE @name TO DISK = @fileName  
 