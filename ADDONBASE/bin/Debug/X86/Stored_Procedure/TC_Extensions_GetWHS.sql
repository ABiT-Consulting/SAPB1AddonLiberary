 procedure TC_Extensions_GetWHS
@itemcode as nvarchar(max),
@whsCode as nvarchar(max),
@serialNumber as int =-1,
@batchnumber as nvarchar(max)=''
as 
if(@serialNumber = -1 )
begin
select "OnHand"+"OnOrder"-"IsCommited" from OITW
where "ItemCode" = @itemcode and "WhsCode" = @whsCode;
end;
if(@serialNumber >-1)
begin
SELECT count(*) FROM OSRI T0 WHERE T0.[WhsCode] =@whsCode and "SysSerial"=@serialNumber
and "ItemCode" = @itemcode
end 
