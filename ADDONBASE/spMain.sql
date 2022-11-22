
declare @str NVARCHAR(50) 
IF EXISTS (SELECT * FROM sys.procedures WHERE SCHEMA_ID=SCHEMA_ID('dbo')AND name= N'{1}')
BEGIN
SET @str = 'Alter'
END
ELSE
SET @str = 'Create'


EXEC ( @str+ '{0}'
)