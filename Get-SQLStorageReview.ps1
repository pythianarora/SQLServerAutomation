[CmdletBinding()]
Param(
		[Parameter(Mandatory=$true)]
		[Alias('hostname')]
		[string]$Computer,

		[Parameter(Mandatory=$true)]
		[Alias('instance')]
		[string]$sqlinstance,
		
		[Parameter(Mandatory=$true)]
		[string]$report
	)
#functions for formatting and logging
function Write-Status ($message) {Write-Host ("[$(get-date -Format 'HH:mm:ss')] $message.").PadRight(75) -NoNewline -ForegroundColor Yellow }
function Update-Status ($status = "Success") {Write-Host "[$status]" -ForegroundColor Green}
function Exit-Fail ($message) {
	Write-Host "`nERROR: $message" -ForegroundColor Red
	Write-Host "Result:Failed." -ForegroundColor Red
	exit 0x1
}
function NoExit-Fail ($message) {
	Write-Host "`nERROR: $message" -ForegroundColor Red
	Write-Host "Result:Continue." -ForegroundColor Red
}
#function to gather details for disk space on the computer
function getdiskspaceinfo ($ComputerName, [int]$DriveType = 3){
	BEGIN{}
	PROCESS{
		Write-Status "Gathering Disk Space information for all disks on $ComputerName"
		if (Test-Connection $ComputerName -Quiet -Count 2) {
			try {
				Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=$DriveType" -ComputerName $ComputerName | 
				Select-Object -Property @{n="Hostname";e={$_.__SERVER}},
				@{n="DriveName";e={$_.DeviceID}},
				@{n="FreeSpace(GB)";e={$_.FreeSpace / 1GB -as [int]}},
				@{n="Size(GB)";e={$_.Size / 1GB -as [int]}},
				@{n="Free(%)";e={$_.FreeSpace / $_.Size * 100 -as [int] }}
				Update-Status
			}
			Catch {
				Exit-Fail $($_.Exception.Message)
			}
		}
		else {
			Exit-Fail "Could not connect to $ComputerName"
		}
	}
	END{}
}
#function to gather sql server instance related details
function getsqlserverinfo ($instance)  {
	BEGIN{
		$ErrorActionPreference = "Stop"
		$query = @"
		select 
		@@servername as 'InstanceName', 
		SERVERPROPERTY('MachineName') as 'MachineName',
		Case SERVERPROPERTY('IsClustered') when 1 then 'CLUSTERED' else 'STANDALONE' end as 'ServerType',
		left(@@VERSION, CHARINDEX(' - ',@@version)-1) as 'Version',
		SERVERPROPERTY ('Edition') as 'Edition',
		SERVERPROPERTY('productversion') 'VersionNumber',
		SERVERPROPERTY ('productlevel') as 'ServicePack',
		(select stuff((SELECT ', '+ cast (ROW_NUMBER() Over(order by dbid) as varchar)+'.'+ UPPER(name)
		FROM sys.sysdatabases FOR XML PATH ('')), 1, 1, '') as 'DB') as 'Databases'
"@
	}
	Process{
		try {
			Write-Status "Getting details for SQL Instance($instance)"
			$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
			$SqlConnection.ConnectionString = "Data Source = $instance; Initial Catalog = master; Integrated Security=true;Connect Timeout=30;"	
			$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
			$SqlCmd.CommandText = $query
			$SqlCmd.Connection = $SqlConnection		
			$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
			$SqlAdapter.SelectCommand = $SqlCmd
			$DataSet = New-Object System.Data.DataSet
			$SqlAdapter.Fill($DataSet) | Out-Null
			Update-Status
			return $DataSet.Tables[0]
		}
		catch {
			Exit-Fail $($_.Exception.Message)
		}
	}
	END{
		$SqlConnection.Close()
		$ErrorActionPreference = "Continue"
	}
}
#function to gather disk offset details for all disks
function getdiskoffset ($ComputerName) {
	BEGIN{}
	PROCESS{
		Write-Status "Gathering Disk offset details on $ComputerName"
		if (Test-Connection $ComputerName -Quiet -Count 2) {
			try {
				$drives = Get-WmiObject Win32_DiskDrive -ComputerName $ComputerName
				$s = New-Object System.Management.ManagementObjectSearcher
				$s.Scope = "\\$Computer\root\cimv2"
				$s2 = New-Object System.Management.ManagementObjectSearcher
				$s2.Scope = "\\$Computer\root\cimv2"
				$qPartition = new-object System.Management.RelatedObjectQuery
				$qPartition.RelationshipClass = 'Win32_DiskDriveToDiskPartition'
				$qLogicalDisk = new-object System.Management.RelatedObjectQuery
				$qLogicalDisk.RelationshipClass = 'Win32_LogicalDiskToPartition'
				$drives | Sort-Object DeviceID | % {
					$qPartition.SourceObject = $_
					$s.Query= $qPartition
					$s.Get()| where {$_.Type -ne 'Unknown'} |% {
						$partition = $_;
						$partitionSize = ([math]::round(($($_.Size)/1GB),1))
						$qLogicalDisk.SourceObject = $_
						$s2.Query= $qLogicalDisk.QueryString
						$s2.Get()|% {
							$props =  @{'ComputerName'=$ComputerName;
										'PartitionName' = $($partition.Name);
										'DriveName'=$($_.DeviceID);
										'FileSystem'=$($_.FileSystem);
										'StartingOffset'=$($partition.StartingOffset)}
							$obj = New-Object -TypeName PSObject -Property $props
							return $obj | select ComputerName,`
													DriveName,` 
													PartitionName,` 
													FileSystem,`
													@{n="Offset" ;e={$_.StartingOffset/1024}} 
						}
					}
				}
				Update-Status
			}
			Catch {
				Exit-Fail $($_.Exception.Message)
			}
		}
		else {
			Exit-Fail "Could not connect to $ComputerName"
		}	
	}
	END{}
}
#function to gather disk blocksize details for all disks
function getblocksize($ComputerName){
	BEGIN{}
	PROCESS{
		Write-Status "Gathering Disk Blocksize details on $ComputerName"
		if (Test-Connection $ComputerName -Quiet -Count 2) {
			try {
				$volumes = Get-WmiObject Win32_Volume -ComputerName $Computer
				Update-Status
				return $volumes | Where-Object {$_.DriveLetter -ne $null -and $_.BlockSize -ne $null } | select @{n="DriveName";e={$_.DriveLetter}}, BlockSize
			}
			catch {
				Exit-Fail $($_.Exception.Message)
			}
		}
		else {
			Exit-Fail "Could not connect to $ComputerName"
		}
	}
	END {}
}

#function to put together blocksize and offset 
#Doesn't do info for Disks Mounted as folders
function getdiskalignment ($hostname) {
	BEGIN{}
	PROCESS{
		$outputCollection = @()
		$offset = getdiskoffset -ComputerName $hostname
		$blocksize = getblocksize -ComputerName $hostname
		
		$offset | Foreach-Object {
			$offsetObject = $_
			$blocksizeObject = $blocksize | Where-Object {$_.DriveName -eq $offsetObject.DriveName}
			$outputObject = "" | Select DriveName, PartitionName, FileSystem, BlockSize, Offset
			$outputObject.DriveName = $offsetObject.DriveName
			$outputObject.PartitionName = $offsetObject.PartitionName
			$outputObject.FileSystem = $offsetObject.Filesystem
			$outputObject.BlockSize = $blocksizeObject.BlockSize
			$outputObject.Offset = $offsetObject.Offset
			$outputCollection += $outputObject
		}
		return $outputCollection | Select-Object DriveName, PartitionName, FileSystem, BlockSize, Offset
	}
	END{}
}
#get mdf and ldf placement for System Databases
function getsystemdbfileloc ($instance)  {
	BEGIN{
		$ErrorActionPreference = "Stop"
		$query = @"
		create table #temp_mdfs(dbname sysname, filename varchar(500), dbid int, fileid int)
		create table #temp_ldfs(dbname sysname, filename varchar(500), dbid int, fileid int)
		insert into #temp_mdfs select db_name(dbid) as dbname, filename , dbid, fileid from sys.sysaltfiles where db_name(dbid) is not NULL and (filename like '%mdf' or filename like '%ndf')
		insert into #temp_ldfs select db_name(dbid) as dbname, filename , dbid, fileid from sys.sysaltfiles where db_name(dbid) is not NULL and filename like '%ldf'
		select m.dbname, m.filename as mdf , l.filename as ldf from #temp_mdfs m, #temp_ldfs l where m.dbid = l.dbid and m.dbname in ('master', 'tempdb', 'model','msdb', 'distribution')
		drop table #temp_mdfs
		drop table #temp_ldfs
"@
	}
	Process{
		try {
			Write-Status "Getting Location for System Databases for SQL Instance($instance)"
			$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
			$SqlConnection.ConnectionString = "Data Source = $instance; Initial Catalog = master; Integrated Security=true;Connect Timeout=30;"
			$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
			$SqlCmd.CommandText = $query
			$SqlCmd.Connection = $SqlConnection
			$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
			$SqlAdapter.SelectCommand = $SqlCmd
			$DataSet = New-Object System.Data.DataSet
			$SqlAdapter.Fill($DataSet) | Out-Null	
			Update-Status
			return $DataSet.Tables[0] 
		}
		catch {
			Exit-Fail $($_.Exception.Message)
		}
	}
	END{
		$SqlConnection.Close()
		$ErrorActionPreference = "Continue"
	}
}

function getdefaultfilelocation ($instance)  {
	BEGIN{
		$ErrorActionPreference = "Stop"
		$query = @"
		IF EXISTS(SELECT 1 FROM [master].[sys].[databases] WHERE [name] = 'zzTempDBForDefaultPath')   
		BEGIN  
		DROP DATABASE zzTempDBForDefaultPath   
		END;
		CREATE DATABASE zzTempDBForDefaultPath;
		DECLARE @Default_Data_Path1 VARCHAR(512),   
			@Default_Log_Path2 VARCHAR(512);
		SELECT @Default_Data_Path1 =    
		(   SELECT LEFT(physical_name,LEN(physical_name)-CHARINDEX('\',REVERSE(physical_name))+1) 
		FROM sys.master_files mf   
		INNER JOIN sys.[databases] d   
		ON mf.[database_id] = d.[database_id]   
		WHERE d.[name] = 'zzTempDBForDefaultPath' AND type = 0);
		SELECT @Default_Log_Path2 =    
		(   SELECT LEFT(physical_name,LEN(physical_name)-CHARINDEX('\',REVERSE(physical_name))+1)   
		FROM sys.master_files mf   
		INNER JOIN sys.[databases] d   
		ON mf.[database_id] = d.[database_id]   
		WHERE d.[name] = 'zzTempDBForDefaultPath' AND type = 1);
		IF EXISTS(SELECT 1 FROM [master].[sys].[databases] WHERE [name] = 'zzTempDBForDefaultPath')   
		BEGIN  
		DROP DATABASE zzTempDBForDefaultPath   
		END						
		select @Default_Data_Path1 as 'data', @Default_Log_Path2 as 'log'
"@
	}
	Process{
		try {
			Write-Status "Getting default locations for SQL Instance($instance)"
			$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
			$SqlConnection.ConnectionString = "Data Source = $instance; Initial Catalog = master; Integrated Security=true;Connect Timeout=30;"	
			$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
			$SqlCmd.CommandText = $query
			$SqlCmd.Connection = $SqlConnection		
			$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
			$SqlAdapter.SelectCommand = $SqlCmd
			$DataSet = New-Object System.Data.DataSet
			$SqlAdapter.Fill($DataSet) | Out-Null
			Update-Status
			return $DataSet.Tables[0]
		}
		catch {
			Exit-Fail $($_.Exception.Message)
		}
	}
	END{
		$SqlConnection.Close()
		$ErrorActionPreference = "Continue"
	}
}

#get mdf and ldf placement for User Databases
function getuserdbfileloc ($instance)  {
	BEGIN{
		$ErrorActionPreference = "Stop"
		$query = @"
		create table #temp_mdfu (dbname sysname, filename varchar(500), dbid int, fileid int)
		create table #temp_ldfu (dbname sysname, filename varchar(500), dbid int, fileid int)
		insert into #temp_mdfu select db_name(dbid) as dbname, filename , dbid, fileid from sys.sysaltfiles where db_name(dbid) is not NULL and (filename like '%mdf' or filename like '%ndf')
		insert into #temp_ldfu select db_name(dbid) as dbname, filename , dbid, fileid from sys.sysaltfiles where db_name(dbid) is not NULL and filename like '%ldf'
		select m.dbname, m.filename as mdf , l.filename as ldf  from #temp_mdfu m, #temp_ldfu l where m.dbid = l.dbid and m.dbname not in ('master', 'tempdb', 'model','msdb','distribution')
		drop table #temp_ldfu
		drop table #temp_mdfu
"@
	}
	Process{
		try {
			Write-Status "Getting Location for User Databases for SQL Instance($instance)"
			$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
			$SqlConnection.ConnectionString = "Data Source = $instance; Initial Catalog = master; Integrated Security=true;Connect Timeout=30;"
			$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
			$SqlCmd.CommandText = $query
			$SqlCmd.Connection = $SqlConnection
			$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
			$SqlAdapter.SelectCommand = $SqlCmd
			$DataSet = New-Object System.Data.DataSet
			$SqlAdapter.Fill($DataSet) | Out-Null
			Update-Status
			return $DataSet.Tables[0] 
		}
		catch {
			Exit-Fail $($_.Exception.Message)
		}
	}
	END{
		$SqlConnection.Close()
		$ErrorActionPreference = "Continue"
	}
}
#get auto growth settings for all databases
function getdbgrowthsettings ($instance)  {
	BEGIN{
		$ErrorActionPreference = "Stop"
		$query = @"
		select DB_NAME(mf.database_id) database_name, mf.name logical_name, physical_name,
		CONVERT (DECIMAL (20,2) , (CONVERT(DECIMAL, size)/128)) [file_size_MB]
		, CASE mf.is_percent_growth
		WHEN 1 THEN 'Yes'
		ELSE 'No'
		END AS [is_percent_growth]
		, CASE mf.is_percent_growth
		WHEN 1 THEN CONVERT(VARCHAR, mf.growth) + '%'
		WHEN 0 THEN CONVERT(VARCHAR, mf.growth/128) + ' MB'
		END AS [growth_in_increment_of]
		, CASE mf.is_percent_growth
		WHEN 1 THEN
		CONVERT(DECIMAL(20,2), (((CONVERT(DECIMAL, size)*growth)/100)*8)/1024)
		WHEN 0 THEN
		CONVERT(DECIMAL(20,2), (CONVERT(DECIMAL, growth)/128))
		END AS [next_auto_growth_size_MB]
		, CASE mf.max_size
		WHEN 0 THEN 'No growth is allowed'
		WHEN -1 THEN 'File will grow until the disk is full'
		ELSE CONVERT(VARCHAR, mf.max_size)
		END AS [max_size]
		from sys.master_files mf
"@
	}
	Process{
		try {
			Write-Status "Getting Auto Growth Settings for all Databases for SQL Instance($instance)"
			$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
			$SqlConnection.ConnectionString = "Data Source = $instance; Initial Catalog = master; Integrated Security=true;Connect Timeout=30;"
			$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
			$SqlCmd.CommandText = $query
			$SqlCmd.Connection = $SqlConnection
			$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
			$SqlAdapter.SelectCommand = $SqlCmd
			$DataSet = New-Object System.Data.DataSet
			$SqlAdapter.Fill($DataSet) | Out-Null
			Update-Status
			return $DataSet.Tables[0] 
		}
		catch {
			Exit-Fail $($_.Exception.Message)
		}
	}
	END{
		$SqlConnection.Close()
		$ErrorActionPreference = "Continue"
	}
}
#Get DB file IO latency info
function getfileiolatency ($instance)  {
	BEGIN{
		$ErrorActionPreference = "Stop"
		$query = @"
		SELECT
			DB_NAME ([vfs].[database_id]) AS [DB],
			LEFT ([mf].[physical_name], 2) AS [Drive],
			[mf].[physical_name],
			[ReadLatency] =
				CASE WHEN [num_of_reads] = 0
					THEN 0 ELSE ([io_stall_read_ms] / [num_of_reads]) END,
			[WriteLatency] =
				CASE WHEN [num_of_writes] = 0
					THEN 0 ELSE ([io_stall_write_ms] / [num_of_writes]) END,
			[Latency] =
				CASE WHEN ([num_of_reads] = 0 AND [num_of_writes] = 0)
					THEN 0 ELSE ([io_stall] / ([num_of_reads] + [num_of_writes])) END,
			[AvgBPerRead] =
				CASE WHEN [num_of_reads] = 0
					THEN 0 ELSE ([num_of_bytes_read] / [num_of_reads]) END,
			[AvgBPerWrite] =
				CASE WHEN [num_of_writes] = 0
					THEN 0 ELSE ([num_of_bytes_written] / [num_of_writes]) END,
			[AvgBPerTransfer] =
				CASE WHEN ([num_of_reads] = 0 AND [num_of_writes] = 0)
					THEN 0 ELSE
						(([num_of_bytes_read] + [num_of_bytes_written]) /
						([num_of_reads] + [num_of_writes])) END
		
		FROM
			sys.dm_io_virtual_file_stats (NULL,NULL) AS [vfs]
		JOIN sys.master_files AS [mf]
			ON [vfs].[database_id] = [mf].[database_id]
			AND [vfs].[file_id] = [mf].[file_id]
		ORDER BY [WriteLatency] DESC;
"@
	}
	Process{
		try {
			Write-Status "Getting I\O Latency for database files for SQL Instance($instance)"
			$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
			$SqlConnection.ConnectionString = "Data Source = $instance; Initial Catalog = master; Integrated Security=true;Connect Timeout=30;"
			$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
			$SqlCmd.CommandText = $query
			$SqlCmd.Connection = $SqlConnection
			$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
			$SqlAdapter.SelectCommand = $SqlCmd
			$DataSet = New-Object System.Data.DataSet
			$SqlAdapter.Fill($DataSet) | Out-Null
			Update-Status
			return $DataSet.Tables[0] 
		}
		catch {
			Exit-Fail $($_.Exception.Message)
		}
	}
	END{
		$SqlConnection.Close()
		$ErrorActionPreference = "Continue"
	}
}

function getdbvlfcount ($instance)  {
	BEGIN{
		$ErrorActionPreference = "Stop"
		$query = @"
		--T-SQL Script Credits - https://gallery.technet.microsoft.com/scriptcenter/SQL-Script-to-list-VLF-e6315249
		--variables to hold each 'iteration'  
		declare @query varchar(100)  
		declare @dbname sysname  
		declare @vlfs int  
		--table variable used to 'loop' over databases  
		declare @databases table (dbname sysname)  
		insert into @databases  
		--only choose online databases  
		select name from sys.databases where state = 0  
		--table variable to hold results  
		declare @vlfcounts table  
			(dbname sysname,  
			vlfcount int)  
		--table variable to capture DBCC loginfo output  
		--changes in the output of DBCC loginfo from SQL2012 mean we have to determine the version 
		declare @MajorVersion tinyint  
		set @MajorVersion = LEFT(CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(max)),CHARINDEX('.',CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(max)))-1) 
		if @MajorVersion < 11 -- pre-SQL2012 
		begin 
			declare @dbccloginfo table  
			(  
				fileid smallint,  
				file_size bigint,  
				start_offset bigint,  
				fseqno int,  
				[status] tinyint,  
				parity tinyint,  
				create_lsn numeric(25,0)  
			)  
			while exists(select top 1 dbname from @databases)  
			begin  
				set @dbname = (select top 1 dbname from @databases)  
				set @query = 'dbcc loginfo (' + '''' + @dbname + ''') '  
		
				insert into @dbccloginfo  
				exec (@query)  
		
				set @vlfs = @@rowcount  
		
				insert @vlfcounts  
				values(@dbname, @vlfs)  
		
				delete from @databases where dbname = @dbname  
			end --while 
		end 
		else 
		begin 
			declare @dbccloginfo2012 table  
			(  
				RecoveryUnitId int, 
				fileid smallint,  
				file_size bigint,  
				start_offset bigint,  
				fseqno int,  
				[status] tinyint,  
				parity tinyint,  
				create_lsn numeric(25,0)  
			)  
			while exists(select top 1 dbname from @databases)  
			begin  
				set @dbname = (select top 1 dbname from @databases)  
				set @query = 'dbcc loginfo (' + '''' + @dbname + ''') '  
				insert into @dbccloginfo2012  
				exec (@query)  
				set @vlfs = @@rowcount  
				insert @vlfcounts  
				values(@dbname, @vlfs)  
				delete from @databases where dbname = @dbname  
			end --while 
		end 
		--output the full list  
		select dbname, vlfcount  
		from @vlfcounts  
		order by dbname
"@
	}
	Process{
		try {
			Write-Status "Getting VLF Count for all Databases for SQL Instance($instance)"
			$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
			$SqlConnection.ConnectionString = "Data Source = $instance; Initial Catalog = master; Integrated Security=true;Connect Timeout=30;"
			$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
			$SqlCmd.CommandText = $query
			$SqlCmd.Connection = $SqlConnection
			$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
			$SqlAdapter.SelectCommand = $SqlCmd
			$DataSet = New-Object System.Data.DataSet
			$SqlAdapter.Fill($DataSet) | Out-Null
			Update-Status
			return $DataSet.Tables[0] 
		}
		catch {
			Exit-Fail $($_.Exception.Message)
		}
	}
	END{
		$SqlConnection.Close()
		$ErrorActionPreference = "Continue"
	}
}

function getdbcccheckdblastrun ($instance)  {
	BEGIN{
		$ErrorActionPreference = "Stop"
		$query = @"
		DECLARE @DB nvarchar(max) = NULL;
		DECLARE @Command nvarchar(max);
		DECLARE @ExecCommand nvarchar(max); 
		CREATE TABLE #DBInfoTemp
		(
		ParentObject varchar(255)
		, [Object] varchar(255)
		, Field varchar(255)
		, [Value] varchar(255)
		);
		CREATE TABLE #LastCkTemp
		(
		DatabaseName varchar(255)
		, LastKnownGoodDate varchar(255)
		);
		IF @DB IS NULL
		BEGIN
		SET @Command = N'
		INSERT INTO #DBInfoTemp
		EXEC (''DBCC DBINFO([?]) WITH TABLERESULTS'');'
		END
		ELSE
		BEGIN
		SET @Command = N'
		INSERT INTO #DBInfoTemp
		EXEC (''DBCC DBINFO([' + @DB + ']) WITH TABLERESULTS'');'
		END
		
		SET @ExecCommand = @Command + N'
		INSERT INTO #LastCkTemp
		SELECT 
		MAX(CASE WHEN di.Field = ''dbi_dbname''
		THEN di.Value
		ELSE NULL
		END) AS DatabaseName    
		, MAX(CASE WHEN di.Field = ''dbi_dbccLastKnownGood''
			THEN di.Value
			ELSE NULL
			END) AS LastCheckDBDate
		FROM #DBInfoTemp di
		WHERE 
		di.Field = ''dbi_dbccLastKnownGood''
		OR di.Field = ''dbi_dbname'';
		
		TRUNCATE TABLE #DBInfoTemp;
		';
		IF @DB IS NULL
		BEGIN
		EXEC sp_MSforeachdb @ExecCommand;
		END
		ELSE
		BEGIN
		EXEC (@ExecCommand);
		END 
		SELECT
		ck.DatabaseName
		, ck.LastKnownGoodDate
		, case 
			when ck.LastKnownGoodDate = '1900-01-01 00:00:00.000'  then 'Never'
			when datediff (day, Getdate(), ck.LastKnownGoodDate ) < 7 then 'Less then 1 Week ago'
			when datediff (day, Getdate(), ck.LastKnownGoodDate ) = 7 then '1 Week ago'
			when datediff (day, Getdate(), ck.LastKnownGoodDate ) >= 7 and datediff (day, Getdate(), ck.LastKnownGoodDate ) <= 14 then '2 Weeks ago'
			when datediff (day, Getdate(), ck.LastKnownGoodDate ) >=14 then 'More than 2 Weeks ago'
			end
		as 'LastRun' 
		FROM #LastCkTemp ck;
		DROP TABLE #LastCkTemp, #DBInfoTemp;
"@
	}
	Process{
		try {
			Write-Status "Getting Last Run Check DB Status for all Databases for SQL Instance($instance)"
			$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
			$SqlConnection.ConnectionString = "Data Source = $instance; Initial Catalog = master; Integrated Security=true;Connect Timeout=30;"
			$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
			$SqlCmd.CommandText = $query
			$SqlCmd.Connection = $SqlConnection
			$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
			$SqlAdapter.SelectCommand = $SqlCmd
			$DataSet = New-Object System.Data.DataSet
			$SqlAdapter.Fill($DataSet) | Out-Null
			Update-Status
			return $DataSet.Tables[0] 
		}
		catch {
			Exit-Fail $($_.Exception.Message)
		}
	}
	END{
		$SqlConnection.Close()
		$ErrorActionPreference = "Continue"
	}
}
#Store info in hash to output report
$serverinfo = getsqlserverinfo -instance $sqlinstance
$volumes = getdiskspaceinfo -ComputerName $Computer
$diskalignment = getdiskalignment -hostname $Computer
$sysfileloc = getsystemdbfileloc -instance $sqlinstance
$defaultloc = getdefaultfilelocation -instance $sqlinstance
$userfileloc = getuserdbfileloc -instance $sqlinstance
$growthstas = getdbgrowthsettings -instance $sqlinstance
$iolatency = getfileiolatency -instance $sqlinstance
$vlfs = getdbvlfcount -instance $sqlinstance
$dbcc = getdbcccheckdblastrun -instance $sqlinstance
$datadrive = $($defaultloc | select data).data
$logdrive = $($defaultloc | select log).log
$datadrive = $datadrive.Split(":\")[0]
$logdrive = $logdrive.Split(":\")[0]

#Final Reporting 
$css = @"
<style>
body	{ 
		background-color:white;
    	font-family: Calibri, sans-serif;
    	font-size:11pt;
		color:#333;
	} 
.header {
			font-size: 40px;
    		color: white;
			text-align: center;
    		background: black;
			font-family: Calibri, sans-serif;
			font-weight: bold;
    		width: 100%;
			border: 1px black;
    		margin: 0;}
h1	{font-size:40px;color: black;}
h2	{font-size:30px;color: black;}
h4  {font-size:15px;color: black;}
table {margin-left:50px; table-layout: fixed;}
td {word-wrap:break-word;}
td , th { 
			border:2px solid black;
         	border-collapse:collapse;
			word-wrap:break-word;
		}
th	{
		font-family: Calibri, sans-serif;
		font-size:14pt; color:white; 
		background-color:black;
		font-weight:bold;
	}
table, tr, td, th {padding: 2px; margin: 1px;}
</style>
"@

if($serverinfo){
	$serverinfo = $serverinfo | select MachineName, InstanceName, ServerType, Version, Edition, VersionNumber, ServicePack, Databases
	$htmlreport = $serverinfo | ConvertTo-HTML -PreContent "<h2><center>SQL Server Information</center></h2><br><br>" -Fragment | out-string
}
else{
	$htmlreport = ConvertTo-HTML -PreContent "<h2><center>SQL Server Information</center></h2><br>" -PostContent "<h4><center>Error Encountered in reporting the details. Check the log.</center></h4><br>" -Fragment | out-string
}
if($volumes){
	$volumes = $volumes | select-Object Hostname,@{N="DriveName";E={$_.DriveName}},`
                                         @{N="DiskSpace(GB)";E={$_."Size(GB)"}},`
                                         @{N="FreeSpace(GB)";E={$_."FreeSpace(GB)"}},`
                                         @{N="Free(%)";e={if($_."Free(%)" -lt 10){'<td bgcolor="#FF0000">'+$_."Free(%)" }elseif($_."Free(%)" -lt 60 -and $_."Free(%)" -gt 10 ){'<td bgcolor="#FFFF00">'+$_."Free(%)"}else{'<td bgcolor="#00CC33">'+$_."Free(%)"}}} 
	$htmlreport += $volumes |  ConvertTo-HTML -PreContent "<h2><center>Disk Space Status</center></h2><br><br>" -Fragment | out-string 
}
else{
	$htmlreport += ConvertTo-HTML -PreContent "<h2><center>Disk Space Status</center></h2><br>" -PostContent "<h4><center>Error Encountered in reporting the disks. Check the log.</center></h4><br>" -Fragment | out-string
}
if($diskalignment){
    $diskalignment = $diskalignment | select-object DriveName,`
                                                    PartitionName,`
                                                    FileSystem,`
                                                    @{n="Block Size";e={if($_.BlockSize -eq 65536){'<td bgcolor="#00CC33">'+$_.BlockSize}else{'<td bgcolor="#FF0000">'+$_.BlockSize}}},`
                                                    @{n="Disk Offset";e={if($_.Offset -eq 1024){'<td bgcolor="#00CC33">'+$_.Offset}else{'<td bgcolor="#FF0000">'+$_.Offset}}}
                                                                            
	$htmlreport += $diskalignment | ConvertTo-HTML -PreContent "<h2><center>Storage Physical Properties</center></h2><br><br>" -Fragment | out-string 	
}
else{
	$htmlreport += ConvertTo-HTML -PreContent "<h2><center>Storage Physical Properties</center></h2><br>" -PostContent "<h4><center>Error Encountered in reporting the properties. Check the log.</center></h4><br>" -Fragment | out-string
}
if($sysfileloc){
	$sysfileloc = $sysfileloc | Select-Object @{n="Database";e={$_.dbname}},`
                                              @{n="MDF File Location";e={if ($_.mdf -like "C:\*"){'<td bgcolor="#FF0000">'+$_.mdf} elseif ($_.dbname -like 'tempdb' -and ($($_.mdf).Split(":\")[0] -like $logdrive -or $($_.mdf).Split(":\")[0] -like $datadrive)) {'<td bgcolor="#FF0000">'+$_.mdf} else{'<td bgcolor="#00CC33">'+$_.mdf}}},`
                                              @{n="LDF File Location";e={if ($_.ldf -like "C:\*"){'<td bgcolor="#FF0000">'+$_.ldf} elseif ($_.dbname -like 'tempdb' -and ($($_.ldf).Split(":\")[0] -like $logdrive -or $($_.ldf).Split(":\")[0] -like $datadrive)) {'<td bgcolor="#FF0000">'+$_.ldf} else{'<td bgcolor="#00CC33">'+$_.ldf}}}
	$htmlreport += $sysfileloc | ConvertTo-HTML -PreContent "<h2><center>System Databases File Location</center></h2><br><br>" -Fragment | out-string
}
else{
	$htmlreport += ConvertTo-HTML -PreContent "<h2><center>System Databases File Location</center></h2><br>" -PostContent "<h4><center>Error Encountered in reporting the properties. Check the log.</center></h4><br>" -Fragment | out-string
}
if($defaultloc){
	$defaultloc = $defaultloc | Select-Object @{n="Database File Default Location ";e={if($($_.data).Split(":\")[0] -like $logdrive){'<td bgcolor="#FF0000">'+$_.data} else{'<td bgcolor="#00CC33">'+$_.data}}},`
                                              @{n="Log File Default Location";e={if($($_.log).Split(":\")[0] -like $datadrive) {'<td bgcolor="#FF0000">'+$_.log}else{'<td bgcolor="#00CC33">'+$_.log}}}
	$htmlreport += $defaultloc | ConvertTo-HTML -PreContent "<h2><center>Database Default Locations</center></h2><br><br>" -Fragment | out-string
}
else{
	$htmlreport += ConvertTo-HTML -PreContent "<h2><center>Database Default Locations</center></h2><br>" -PostContent "<h4><center>Error Encountered in reporting the properties. Check the log.</center></h4><br>" -Fragment | out-string
}
if($userfileloc){
    $datadrive = $datadrive + ":"
    $logdrive = $logdrive + ":"
	$userfileloc = $userfileloc | Select-Object @{n="Database";e={$_.dbname}},`
                                                 @{n="MDF File Location";e={if($_.mdf -like "$logdrive\*"){'<td bgcolor="#FF0000">'+$_.mdf} else{'<td bgcolor="#00CC33">'+$_.mdf}}},`
                                                 @{n="LDF File Location";e={if ($_.ldf -like "$datadrive\*") {'<td bgcolor="#FF0000">'+$_.ldf}else{'<td bgcolor="#00CC33">'+$_.ldf}}}
	$htmlreport += $userfileloc | ConvertTo-HTML -PreContent "<h2><center>User Database(s) File Location</center></h2><br><br>" -Fragment | out-string
}
else{
	$htmlreport += ConvertTo-HTML -PreContent "<h2><center>User Database(s) File Location</center></h2><br>" -PostContent "<h4><center>Error Encountered in reporting the properties. Check the log.</center></h4><br>" -Fragment | out-string
}
if($growthstas){
	$growthstas = $growthstas | select-object @{n="Database Name";e={$_.database_name}},` 
                                                @{n="Logical File Name";e={$_.logical_name}},`
                                                @{n="Physical File Name";e={$_.physical_name}},`
                                                @{n="File Size (MB)";e={$_.file_size_MB}},`
                                                @{n="Is Percent Growth?";e={$_.is_percent_growth}},`
                                                @{n="Growth Increment Factor";e={if($_.growth_in_increment_of -like '*%') {'<td bgcolor="#FF0000">'+$_.growth_in_increment_of} else {'<td bgcolor="#00CC33">'+$_.growth_in_increment_of}}},`
                                                @{n="Next Auto Growth Size (MB)";e={$_.next_auto_growth_size_MB}},`
                                                @{n="Max File Size";e={$_.max_size}}
    $htmlreport += $growthstas | ConvertTo-HTML -PreContent "<h2><center>Database Auto-Growth Stats</center></h2><br><br>" -Fragment | out-string
}
else{
	$htmlreport += ConvertTo-HTML -PreContent "<h2><center>Database Auto-Growth Stats</center></h2><br>" -PostContent "<h4><center>Error Encountered in reporting the properties. Check the log.</center></h4><br>" -Fragment | out-string
}
if($iolatency){
	$iolatency = $iolatency | Select-Object @{n="Database";e={$_.DB}}, Drive, @{n="Physical Location";e={$_.physical_name}}, ReadLatency, WriteLatency, @{n="Latency";e={if($_.Latency -lt 10 ){'<td bgcolor="#00CC33">'+$_.latency} else{'<td bgcolor="#FF0000">'+$_.latency}}}, AvgBPerRead, AvgBPerWrite, AvgBPerTransfer
	$htmlreport += $iolatency | ConvertTo-HTML -PreContent "<h2><center>Database File IO Latency</center></h2><br><br>" -Fragment | out-string
}
else{
	$htmlreport += ConvertTo-HTML -PreContent "<h2><center>Database File IO Latency</center></h2><br>" -PostContent "<h4><center>Error Encountered in reporting the properties. Check the log.</center></h4><br>" -Fragment | out-string
}
if($vlfs){
	$vlfs = $vlfs | select @{n="Database Name";e={$_.dbname}},` 
                            @{n="VLF Count";e={if($_.vlfcount -lt 50) {'<td bgcolor="#00CC33">'+$_.vlfcount} else {'<td bgcolor="#FF0000">'+$_.vlfcount}}}
    $htmlreport += $vlfs | ConvertTo-HTML -PreContent "<h2><center>Check Database VLF Count</center></h2><br><br>" -Fragment | out-string
}
else{
	$htmlreport += ConvertTo-HTML -PreContent "<h2><center>Check Database VLF Count</center></h2><br>" -PostContent "<h4><center>Error Encountered in reporting the properties. Check the log.</center></h4><br>" -Fragment | out-string
}
if($dbcc){
	$dbcc = $dbcc | Select DatabaseName, @{n="LastKnownGoodDate";e={if($_.LastKnownGoodDate -like '1900-01-01 00:00:00.000') {'<td bgcolor="#FF0000">'+"Never"} else {$_.LastKnownGoodDate}}},
                            @{n="LastRun";e={if($_.LastRun -like '*1*week*'){'<td bgcolor="#00CC33">'+$_.LastRun}elseif ($_.LastRun -like '*2*week*'){'<td bgcolor="#FFFF00">'+$_.LastRun} else{'<td bgcolor="#FF0000">'+$_.LastRun}}} 
	$htmlreport += $dbcc | ConvertTo-HTML -PreContent "<h2><center>Check Last Run DBCC CHECKDB</center></h2><br><br>" -Fragment | out-string
}
else{
	$htmlreport += ConvertTo-HTML -PreContent "<h2><center>Check Last Run DBCC CHECKDB</center></h2><br>" -PostContent "<h4><center>Error Encountered in reporting the properties. Check the log.</center></h4><br>" -Fragment | out-string
}

$htmlreport = $htmlreport.Replace('<td>False</td>','<td bgcolor="#FF0000">False</td>')
$htmlreport = $htmlreport.Replace('<td>True</td>','<td bgcolor="#00CC33">True</td>')
$htmlreport = $htmlreport.Replace('<td>Not Correct</td>','<td bgcolor="#FF0000">Not Correct</td>')
$htmlreport = $htmlreport.Replace('<td>Correct</td>','<td bgcolor="#00CC33">Correct</td>')
$htmlreport = $htmlreport.Replace('<td>Yes</td>','<td bgcolor="#FF0000">Yes</td>')
$htmlreport = $htmlreport.Replace('<td>No</td>','<td bgcolor="#00CC33">No</td>')
$htmlreport = $htmlreport.Replace('<td>&lt;td bgcolor=&quot;#FFFF00&quot;&gt;','<td bgcolor="#FFFF00">')
$htmlreport = $htmlreport.Replace('<td>&lt;td bgcolor=&quot;#FFFF00&quot;&gt;','<td bgcolor="#FFFF00">')
$htmlreport = $htmlreport.Replace('<td>&lt;td bgcolor=&quot;#FF0000&quot;&gt;','<td bgcolor="#FF0000">')
$htmlreport = $htmlreport.Replace('<td>&lt;td bgcolor=&quot;#00CC33&quot;&gt;','<td bgcolor="#00CC33">')

try {
	Write-Status "Generating HTML Report at $report"
	$body = convertto-html -Head $css -PostContent "$htmlreport <br/><br/><h1 class = `"header`">END OF REPORT</h1>" -body "<h1 class = `"header`">SQL SERVER STORAGE REVIEW</h1><center><h4>Report Generated on $(Get-date -DisplayHint date -Format g)</h4></center>" -Title "SQL SERVER HEALTH CHECK REPORT" | out-string
	$body = $body.Replace("<table>`r`n</table>","")
	$body | Out-File "$report"
	Update-Status
}
Catch {
	Exit-Fail $($_.Exception.Message)
}

<#
.SYNOPSIS
    The Script retrives the best practise storage information for servers with SQL Serveron it and put's all that info in a html document.
.DESCRIPTION
    The Following information structures can be found in the report.
1. Server Information
2. Disk Space Status (Used Space, Free Space etc.)
3. Disk partition alignment
4. System data base file location
5. User data base file location
6. File Autogrowth Staistics
7. Most Common IO related wait types
8. IO Latency for DB files
9. VLF count for all the databases
10. Last run DBCC Check DB Status.
.PARAMETER Computer
Specify the Computer\Machine\Sever Name here to make IO analysis for the Server. One hostname per run.
For Example: Get-SQLStorageReview -computer localhost ......
.PARAMETER instance
Specify the SQL instance Name to make IO analysis for the Server. One SQL instance per run.
For Example: Get-SQLStorageReview ...... -instance .\jupiter ......
.PARAMETER report
Specify the location for the output report
For Example: Get-SQLStorageReview ...... C:\temp\iotest.html
.EXAMPLE
Get-SQLStorageReview -computer localhost -instance .\jupiter -report C:\temp\StorageReview-Jupiter.html
This will generate a html color coded report for localhost with sql instance jupiter
and will dump the report @ location - C:\temp\StorageReview-Jupiter.html
#>