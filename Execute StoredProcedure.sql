USE [YourDBName]
GO

DECLARE @RC int
DECLARE @verbose bit = 0 -- 0 is debug mode, 1 is execute the command
DECLARE @companyName nvarchar(50) = N'' -- Cronus -- blank for all companies
DECLARE @tblFilter nvarchar(max) = N''  -- CONCAT('AND [Object ID] BETWEEN',SPACE(1),36,SPACE(1),'AND',SPACE(1),37) -- OR -- CONCAT('AND [Object Name] = ','''','Customer','''')
DECLARE @fldLength int = 30 -- Most text field length is 30
DECLARE @dropLogTable bit = 1 -- 0 is not drop the UpdateLog table, 1 is drop the UpdateLog table (usefule if you are using table by table)
DECLARE @maxDOP tinyint = 0

-- TODO: Set parameter values here.
EXECUTE @RC = [dbo].[RemoveCrLf] 
	@verbose
	,@companyName
	,@tblFilter
	,@fldLength
	,@dropLogTable
	,@maxDOP
GO