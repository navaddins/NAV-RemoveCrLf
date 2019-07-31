-- ================================================
-- Template generated from Template Explorer using:
-- Create Procedure (New Menu).SQL
--
-- This block of comments will not be included in
-- the definition of the procedure.
-- ================================================
SET ANSI_NULLS ON
SET QUOTED_IDENTIFIER ON
GO

IF (OBJECT_ID('CommandExecute') IS NOT NULL)
    DROP PROCEDURE CommandExecute
GO

CREATE PROCEDURE [dbo].[CommandExecute] @varDBName nvarchar(250),
@varSQL nvarchar(max),
@varUseBeginCommit bit = 1,
@varResult bit = 0 OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    SET ANSI_PADDING ON;
    SET ANSI_WARNINGS ON;
    SET ARITHABORT ON;
    SET CONCAT_NULL_YIELDS_NULL ON;
    SET QUOTED_IDENTIFIER ON;
    SET NUMERIC_ROUNDABORT OFF;
    SET LANGUAGE us_english;
    SET DATEFORMAT mdy;
    SET DATEFIRST 7;
    BEGIN TRY
		IF (@varUseBeginCommit = 1)
        	BEGIN TRANSACTION;
			
            IF (LEN(@varDBName) <> 0)
                EXECUTE (N'USE QUOTENAME(PARSENAME(' + @varDBName + ',1)); EXEC sp_executesql N' + @varSQL);
            ELSE
                EXEC sp_executesql @varSQL;
        
		IF (@varUseBeginCommit = 1)
			COMMIT TRANSACTION;
        SET @varResult = 1;
    END TRY
    BEGIN CATCH
        IF (@varUseBeginCommit = 1)
			IF (@@TRANCOUNT > 0)
				ROLLBACK TRANSACTION;
        THROW;
    END CATCH
END
    RETURN @varResult;
GO

-- ================================================
-- Template generated from Template Explorer using:
-- Create Procedure (New Menu).SQL
--
-- This block of comments will not be included in
-- the definition of the procedure.
-- ================================================
-- =============================================
-- Author:		NavAddIns
-- Create date: 2018-08-08
-- Description:	Remove CrLf from existing data for (Navision)
-- Remarks: 
--			Table 
--				[UpdateLog] will create for log. It will delete base on parameter after process is done.
--			Stored Procedures 
--				[RemoveCrLf] is created
--				[CommandExecute] is created
-- =============================================
SET ANSI_NULLS ON
SET QUOTED_IDENTIFIER ON
GO

IF (OBJECT_ID('RemoveCrLf') IS NOT NULL)
    DROP PROCEDURE RemoveCrLf
GO

CREATE PROCEDURE [dbo].[RemoveCrLf] @verbose bit = 0,
@companyName nvarchar(50) = N'',
@tblFilter nvarchar(max) = N'',
@fldLength int = 30,
@dropLogTable bit = 1,
@maxDOP tinyint = 0
AS
BEGIN
    SET NOCOUNT ON;
    SET ANSI_PADDING ON;
    SET ANSI_WARNINGS ON;
    SET ARITHABORT ON;
    SET CONCAT_NULL_YIELDS_NULL ON;
    SET NUMERIC_ROUNDABORT OFF;
    SET LANGUAGE us_english;
    SET DATEFORMAT mdy;
    SET DATEFIRST 7;

    DECLARE @company nvarchar(250) = N'';
    DECLARE @databaseversionno int = 0;
    DECLARE @datafilegroup nvarchar(max) = N'';

    DECLARE @sql nvarchar(max) = N'';
    DECLARE @table nvarchar(250) = N'';
    DECLARE @columnList nvarchar(max) = N'';
    DECLARE @columnListNaked nvarchar(max) = N'';
    DECLARE @columnName nvarchar(100) = N'';
    DECLARE @columnType int;
    DECLARE @columnTypeNameVarchar nvarchar(8) = N'';
	DECLARE @columnTypeNameNVarchar nvarchar(8) = N'';
    DECLARE @columnId AS int = 0;
	DECLARE @maxColumnId AS int = 0;
	DECLARE @maxLength int;

    DECLARE @notinStart int = 2000000001; -- System Table Start
    DECLARE @notinEnd int = 2000000300; -- System Table End
    DECLARE @isFirstLoop bit = 1; -- Do not change this value

    DECLARE @dbName nvarchar(max) = DB_NAME();
    DECLARE @tableNo int = 0;
    DECLARE @tableName nvarchar(250) = N'';
    DECLARE @patternTableName nvarchar(30) = QUOTENAME(N'\.|\/|\[|\]|\"');
    DECLARE @replaceTableName nvarchar(10) = N'_';
    DECLARE @rowCount AS int = 0;
	DECLARE @rowNumber as int = 0;	
    DECLARE @Result bit = 0;

    DECLARE @codepagefrom int;
    DECLARE @codepageto int;

    --- This is hardcode do not change ---
    SET @columnTypeNameVarchar = 'varchar'; -- Use lowercase
	SET @columnTypeNameNVarchar = 'nvarchar'; -- Use lowercase
    DECLARE @StartDate datetime
    DECLARE @StartDatePerComp datetime
    DECLARE @EndDate datetime
    DECLARE @Duration nvarchar(20) = N'';
    DECLARE @compatibilityLevel int = 0;
    DECLARE @PrintMsg nvarchar(max) = N''
    DECLARE @MsgString nvarchar(max) = N'This is an informational message only. No user action is required.<<%s>>'
    DECLARE @MsgPrevString nvarchar(max) = N'This is an informational message only. No user action is required.<<Step No. %d. is already done.>>'
	DECLARE @WaitForDelay nvarchar(20) = '00:00:00.03'
	DECLARE @skipEmptyTable bit = 1
    SELECT
        @compatibilityLevel = [compatibility_level]
    FROM sys.databases
    WHERE sys.databases.[name] = @dbName;

    SELECT
        @datafilegroup = QUOTENAME([name])
    FROM sys.filegroups
    WHERE is_default = 1;

    SELECT
        @databaseversionno = [databaseversionno]
    FROM [dbo].[$ndo$dbproperty];

    -- Print Info.
    RAISERROR ('', 10, 1) WITH NOWAIT;
    RAISERROR ('-- Source DBName :%s', 10, 1, @dbName) WITH NOWAIT;
    RAISERROR ('-- Data File Group :%s', 10, 1, @datafilegroup) WITH NOWAIT;
    RAISERROR ('-- Filtered Companay :%s', 10, 1, @companyName) WITH NOWAIT;
	RAISERROR ('-- Filtered Table :%s', 10, 1, @tblFilter) WITH NOWAIT;
    IF (@verbose = 1)
        RAISERROR ('-- Debugging mode :%s', 10, 1, 'False') WITH NOWAIT;
    ELSE
        RAISERROR ('-- Debugging mode :%s', 10, 1, 'True') WITH NOWAIT;        
    
	IF (@dropLogTable = 1)
        RAISERROR ('-- Drop the UpdateLog Table :%s', 10, 1, 'True') WITH NOWAIT;
    ELSE
        RAISERROR ('-- Drop the UpdateLog Table :%s', 10, 1, 'False') WITH NOWAIT;
	
	RAISERROR ('-- SQL Version :%s', 10, 1, @@VERSION) WITH NOWAIT;
	RAISERROR ('-- Maximum Degree of Parallelism :%d', 10, 1, @maxDOP) WITH NOWAIT;
    
	IF OBJECT_ID('CommandExecute') IS NULL
    BEGIN
        RAISERROR ('CommandExecute Stored Procedure is missing.', 16, 1);
        RETURN;
    END;

    IF (@verbose = 1)
        IF OBJECT_ID('[dbo].UpdateLog') IS NULL
        BEGIN
            CREATE TABLE [dbo].[UpdateLog] (
                [TableNo] [int] NOT NULL,
                [TableName] [nvarchar](250) NOT NULL,
                [TotalRecord] [int] NULL,
                [StepNo] [smallint] IDENTITY(1,1) NOT NULL,
                [ExecuteCmd] [nvarchar](MAX) NULL,
                [ProcessedTime] [time] NOT NULL
                CONSTRAINT [PK_UpdateLog] PRIMARY KEY CLUSTERED
                (
                [TableName] ASC,
                [StepNo] ASC
                ) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
            ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
            ALTER TABLE [dbo].[UpdateLog] ADD CONSTRAINT [DF_UpdateLog_TableNo] DEFAULT ((0)) FOR [TableNo]
            ALTER TABLE [dbo].[UpdateLog] ADD CONSTRAINT [DF_UpdateLog_TotalRecord] DEFAULT ((0)) FOR [TotalRecord]
        END

    IF (@databaseversionno > 0)
    BEGIN
    BEGIN TRY
        --Disable the Auto Update Statistics
        IF (@verbose = 1)
            EXECUTE [dbo].[CommandExecute] N'',N'ALTER DATABASE CURRENT SET AUTO_UPDATE_STATISTICS OFF;',0,@Result OUTPUT;

        SET @StartDate = GETDATE()
        SET @PrintMsg = '-- Start Time :' + CONVERT(varchar(20), @StartDate, 120);
        RAISERROR (N'%s', 10, 1, @PrintMsg) WITH NOWAIT;
		RAISERROR (N'%s', 10, 1, '') WITH NOWAIT;
		WAITFOR DELAY @WaitForDelay

        SET @PrintMsg = '-- Collection the table information';
        RAISERROR (N'%s', 10, 1, @PrintMsg) WITH NOWAIT;
		WAITFOR DELAY @WaitForDelay
        
        -- Use temp for combine the table lists
        IF OBJECT_ID('tempdb..#tmpSystemObject') IS NOT NULL
            DROP TABLE #tmpSystemObject		
		SELECT 
            SUM(sPTN.[rows]) AS [SysObjs_ Record Count],
            (CASE
                WHEN (CHARINDEX('$', sOBJ.name) = 0) THEN ''
                ELSE (
                    CASE
                        WHEN (LEFT(sOBJ.name, CHARINDEX('$', sOBJ.name) - 1) <> '') THEN LEFT(sOBJ.name, CHARINDEX('$', sOBJ.name) - 1)
                        ELSE ''
                    END
                    )
            END) AS [SysObjs_ Company Name],
            (CASE
                WHEN (CHARINDEX('$', sOBJ.name) = 0) THEN sOBJ.name
                ELSE (
                    CASE
                        WHEN (LEFT(sOBJ.name, CHARINDEX('$', sOBJ.name) - 1) <> '') THEN RIGHT(sOBJ.name, LEN(sOBJ.name) - CHARINDEX('$', sOBJ.name))
                        ELSE sOBJ.name
                    END
                    )
            END) AS [SysObjs_ Name],
            sOBJ.name AS [SysObjs_ Table Name] INTO #tmpSystemObject
        FROM sys.objects AS sOBJ
        INNER JOIN sys.partitions AS sPTN
            ON sOBJ.object_id = sPTN.object_id
        WHERE sOBJ.type = 'U'
        AND sOBJ.is_ms_shipped = 0x0
        AND index_id < 2 -- 0:Heap, 1:Clustered
        GROUP BY sOBJ.name
		
		IF OBJECT_ID('tempdb..#tmpObjects') IS NOT NULL
			DROP TABLE #tmpObjects       
		SELECT
            Objs.[ID] AS [Object ID],
            Objs.[Name] AS [Object Name],
            [SysObjs_ Table Name] AS [Table Name],
            [SysObjs_ Record Count] AS [Record Count],
            [SysObjs_ Company Name] AS [Company Name] INTO #tmpObjects
        FROM #tmpSystemObject AS [tmpSysObjects]
        INNER JOIN [Object] AS Objs
            ON [SysObjs_ Name] = master.[dbo].ReplaceString(Objs.[Name], @patternTableName, @replaceTableName, 2)
            AND [SysObjs_ Company Name] = master.[dbo].ReplaceString(Objs.[Company Name], @patternTableName, @replaceTableName, 2)
        WHERE (Objs.[Type] = 0) AND (NOT (Objs.[ID] BETWEEN @notinStart AND @notinEnd))
		
		-- Show Object List Before start
		SELECT * FROM #tmpObjects ORDER BY [Object ID]
        
		DECLARE cmp_cursor CURSOR FAST_FORWARD FOR
        SELECT
            [Name]
        FROM [dbo].Company
        WHERE CASE
            WHEN ((@companyName IS NULL) OR
                (LEN(@companyName) = 0)) THEN ''
            ELSE [Name]
        END =
             CASE
                 WHEN ((@companyName IS NULL) OR
                     (LEN(@companyName) = 0)) THEN ''
                 ELSE @companyName
             END
        ORDER BY [dbo].Company.[Name]
        OPEN cmp_cursor
        FETCH NEXT FROM cmp_cursor INTO @company
        WHILE (@@FETCH_STATUS <> -1)
        BEGIN			
            SET @StartDatePerComp = GETDATE();
            SELECT
                @company = master.[dbo].ReplaceString(@company, @patternTableName, @replaceTableName, 2);            			
			DECLARE @tbl_cursor CURSOR
            SET @sql = CONCAT(
				N'SET @tbl_cursor = CURSOR FAST_FORWARD FOR',SPACE(1),
				'SELECT
					[Object ID],
					[Table Name],
					[Record Count]
				FROM #tmpObjects
				WHERE NOT ([Object ID] BETWEEN @notinStart AND @notinEnd)
				AND (
				[Record Count] >
								CASE
									WHEN (@skipEmptyTable = 1) THEN 0
									ELSE -1
								END
				)
				AND (
				([Company Name] =
								 CASE
									 WHEN (@isFirstLoop = 1) THEN ''''
									 ELSE @company
								 END
				)
				OR ([Company Name] = @company)
				)', CHAR(10), @tblFilter, SPACE(1),
				'ORDER BY [Company Name], [Object Name];',CHAR(10),
				'OPEN @tbl_cursor;')
			EXEC sp_executesql @sql,
				N'@tblFilter nvarchar(max), @notinStart int, @notinEnd int, @skipEmptyTable bit,@isFirstLoop bit, 
				@company nvarchar(50), @tableNo int, @table nvarchar(250), @rowCount int, @tbl_cursor CURSOR OUTPUT',
			@tblFilter = @tblFilter,
			@notinStart = @notinStart,
			@notinEnd = @notinEnd,
			@skipEmptyTable = @skipEmptyTable,
			@isFirstLoop = @isFirstLoop,
			@company = @company,
			@tableNo = @tableNo,
			@table = @table,
			@rowCount = @rowCount,
			@tbl_cursor = @tbl_cursor OUTPUT
            FETCH NEXT FROM @tbl_cursor INTO @tableNo, @table, @rowCount
            WHILE (@@FETCH_STATUS <> -1)
            BEGIN
                RAISERROR ('', 10, 1) WITH NOWAIT
				SET @isFirstLoop = 0;
                SET @tableName = @table;
                                                
                DECLARE @hasTable bit = 0;
                SELECT
                    @hasTable = COUNT(*)
                FROM sys.tables
                WHERE sys.tables.[name] = PARSENAME(@tableName, 1);
                IF (@hasTable = 0)
                BEGIN
                    SET @PrintMsg = '-- Table' + SPACE(1) + QUOTENAME(@tableName) + SPACE(1) + 'does not exist in' + SPACE(1) + DB_NAME();
                    RAISERROR (@MsgString, 10, 1, @PrintMsg) WITH NOWAIT
                END
                ELSE
                BEGIN
                    IF (@rowCount = 0)                        
                    BEGIN
                        SET @PrintMsg = '-- Table' + SPACE(1) + QUOTENAME(@tableName) + SPACE(1) + 'record is zero. Nothing to do for' + SPACE(1) + QUOTENAME(@tableName);
                        RAISERROR (@MsgString, 10, 1, @PrintMsg) WITH NOWAIT
                    END
                    ELSE
                    BEGIN
                        SET @tableName = QUOTENAME(@tableName);                        
                        SET @PrintMsg = '-- Total Records (' + CAST(@rowCount AS nvarchar(max)) + ') of' + SPACE(1) + @tableName;
                        RAISERROR (N'%s', 10, 1, @PrintMsg) WITH NOWAIT;
                        -- Create column list
                        SET @columnList = N'';
                        SET @columnListNaked = '';
						                        
						SELECT @maxColumnId = max(column_id) 
							FROM sys.columns
							WHERE sys.columns.[object_id] = OBJECT_ID(PARSENAME(@tableName, 1))
							AND sys.columns.[system_type_id] <> 189 -- Leave out timestamp column
							AND sys.columns.[max_length] >= @fldLength
							AND [system_type_id] IN (TYPE_ID(@columnTypeNameVarchar), TYPE_ID(@columnTypeNameNVarchar))
							
						DECLARE col_cursor CURSOR FAST_FORWARD FOR
                        SELECT
							column_id,
                            [name],
                            [system_type_id],
                            [max_length]
                        FROM sys.columns
                        WHERE sys.columns.[object_id] = OBJECT_ID(PARSENAME(@tableName, 1))
                        AND sys.columns.[system_type_id] <> 189 -- Leave out timestamp column
						AND sys.columns.[max_length] >= @fldLength
						AND [system_type_id] IN (TYPE_ID(@columnTypeNameVarchar), TYPE_ID(@columnTypeNameNVarchar))
                        ORDER BY column_id
						OPEN col_cursor
                        FETCH NEXT FROM col_cursor INTO @columnId,@columnName, @columnType, @maxLength;
                        WHILE (@@FETCH_STATUS <> -1)
                        BEGIN
                            SET @columnName = QUOTENAME(@columnName);
                            IF ((@columnType = TYPE_ID(@columnTypeNameVarchar)) OR (@columnType = TYPE_ID(@columnTypeNameNVarchar)))
                            BEGIN
                                SET @columnList = CONCAT('master.[dbo].ReplaceString(',@columnName,',','''','[\t|\n|\r|\s]','''',',',''' ''',',',2,')')
                            END

							SET @columnListNaked = CONCAT(@columnListNaked,@columnName,SPACE(1),'=',SPACE(1),@columnList);
							IF (@columnId < @maxColumnId)
								SET @columnListNaked = CONCAT(@columnListNaked,',');							
							SET @columnListNaked = CONCAT(@columnListNaked,CHAR(10));
                            
							FETCH NEXT FROM col_cursor INTO @columnId,@columnName, @columnType, @maxLength;							
                        END
                        CLOSE col_cursor
                        DEALLOCATE col_cursor

                        IF (LEN(@columnList) = 0)
                        BEGIN
                            SET @PrintMsg = '-- Nothing to do for' + SPACE(1) + @tableName;
                            RAISERROR (N'%s', 10, 1, @PrintMsg) WITH NOWAIT;
							WAITFOR DELAY @WaitForDelay
                        END
						ELSE BEGIN
							SET @sql = CONCAT('SET STATISTICS TIME ON;',SPACE(1),CHAR(10),
										'UPDATE [dbo].',@tableName,SPACE(1),'SET',SPACE(1),CHAR(10),
										@columnListNaked,'OPTION',SPACE(1),'(MAXDOP',SPACE(1),CAST(@maxDOP AS nvarchar(1)),');',SPACE(1),CHAR(10),
										'SET STATISTICS TIME OFF;');
						
							WAITFOR DELAY @WaitForDelay	
							SET @PrintMsg = @sql;
							RAISERROR (N'%s', 10, 1, @PrintMsg) WITH NOWAIT;                        
							IF (@verbose = 1)
							BEGIN
								SET @PrintMsg = '--' + SPACE(1) + @tableName + SPACE(1) + 'is already updated';
								SELECT @Result = COUNT(1) FROM [dbo].UpdateLog WHERE [dbo].UpdateLog.[TableName] = @tableName
								IF (@Result = 0)
									EXECUTE [dbo].[CommandExecute] N'', @sql, 1,@Result OUTPUT;
								ELSE
								BEGIN
									SET @Result = 0
									RAISERROR (N'%s', 10, 1, '') WITH NOWAIT;
									RAISERROR (N'%s', 10, 1, @PrintMsg) WITH NOWAIT;
								END;
								SET @sql = CONCAT('INSERT INTO [dbo].UpdateLog(TableName, TotalRecord, ExecuteCmd, ProcessedTime)',SPACE(1),
												  'VALUES',SPACE(1),'(','''',@tableName,'''',',',@rowCount, ',', '''', REPLACE(@sql,'''',''''''), '''', ',', '''', GETDATE(), '''',');')
                                IF (@Result <> 0)
                                    EXECUTE [dbo].[CommandExecute] N'',@sql,1,@Result OUTPUT;
							END
						END
					END
				END				
				FETCH NEXT FROM @tbl_cursor INTO @tableNo, @table, @rowCount
            END
            CLOSE @tbl_cursor
            DEALLOCATE @tbl_cursor

            RAISERROR ('', 10, 1) WITH NOWAIT;
            SET @PrintMsg = '-- Remove CrLf is done for' + SPACE(1) + @company + '.'
            RAISERROR (N'%s', 10, 1, @PrintMsg) WITH NOWAIT;
            SET @EndDate = GETDATE();
            SELECT
                @Duration = CONVERT(varchar(5), DATEDIFF(s, @StartDatePerComp, @EndDate) / 3600) + ':'
                + CONVERT(varchar(5), DATEDIFF(s, @StartDatePerComp, @EndDate) % 3600 / 60) + ':'
                + CONVERT(varchar(5), (DATEDIFF(s, @StartDatePerComp, @EndDate) % 60))

            SET @PrintMsg = '-- Start Time :' + CONVERT(varchar(20), @StartDatePerComp, 120);
            RAISERROR (N'%s', 10, 1, @PrintMsg) WITH NOWAIT;
            SET @PrintMsg = '-- End Time :' + CONVERT(varchar(20), @EndDate, 120);
            RAISERROR (N'%s', 10, 1, @PrintMsg) WITH NOWAIT;
            SET @PrintMsg = '-- Total Duration :' + SPACE(1) + @Duration
            RAISERROR (N'%s', 10, 1, @PrintMsg) WITH NOWAIT;

            RAISERROR ('------------------------------------------------------------------', 10, 1) WITH NOWAIT;

            FETCH NEXT FROM cmp_cursor INTO @company
        END
        CLOSE cmp_cursor
        DEALLOCATE cmp_cursor

        SET @EndDate = GETDATE();
        SELECT
            @Duration = CONVERT(varchar(5), DATEDIFF(s, @StartDate, @EndDate) / 3600) + ':'
            + CONVERT(varchar(5), DATEDIFF(s, @StartDate, @EndDate) % 3600 / 60) + ':'
            + CONVERT(varchar(5), (DATEDIFF(s, @StartDate, @EndDate) % 60))

        RAISERROR ('', 10, 1) WITH NOWAIT;
        SET @PrintMsg = '-- Start Time :' + CONVERT(varchar(20), @StartDate, 120);
        RAISERROR (@PrintMsg, 10, 1) WITH NOWAIT;
        SET @PrintMsg = '-- End Time :' + CONVERT(varchar(20), @EndDate, 120);
        RAISERROR (@PrintMsg, 10, 1) WITH NOWAIT;
        SET @PrintMsg = '-- Overall Duration :' + SPACE(1) + @Duration
        RAISERROR (@PrintMsg, 10, 1) WITH NOWAIT;

		IF (@verbose = 1)
        BEGIN
            SELECT * FROM [dbo].UpdateLog;
			IF (@dropLogTable = 1)
				IF OBJECT_ID('[dbo].UpdateLog') IS NOT NULL
					DROP TABLE [dbo].UpdateLog;
        END;
        --Enable the Auto Update Statistics
        IF (@verbose = 1)
		    EXECUTE [dbo].[CommandExecute] N'',N'ALTER DATABASE CURRENT SET AUTO_UPDATE_STATISTICS ON;',0,@Result OUTPUT;
    END TRY
    BEGIN CATCH
		--Enable the Auto Update Statistics
        IF (@verbose = 1)
            EXECUTE [dbo].[CommandExecute] N'',N'ALTER DATABASE CURRENT SET AUTO_UPDATE_STATISTICS ON;',0,@Result OUTPUT;
		THROW;
    END CATCH		
    END
END
GO
