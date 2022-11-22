create FUNCTION TC_Functions_fnGetSAPTime
(
	@tm SMALLINT
)
RETURNS NCHAR(10)
AS
BEGIN
	DECLARE @tmOut NCHAR(10)
	
	SELECT @tmOut =  CASE WHEN LEN(@tm) = 4 THEN
		LEFT(@tm,2) + ':' + RIGHT(@tm,2)
		WHEN LEN(@tm) = 3 THEN
		'0'+LEFT(@tm,1) + ':' + RIGHT(@tm,2)
		WHEN LEN(@tm) = 0 THEN
		 '00:00'
		END
  
        
	RETURN @tmOut

END