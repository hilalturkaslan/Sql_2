EXEC sp_enum_oledb_providers;


SELECT * 
FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0', 
                'Excel 12.0;Database=C:\path\to\yourfile.xlsx;HDR=YES', 
                'SELECT * FROM [Sheet1$]');



-- Geliþmiþ seçenekleri aç
EXEC sp_configure 'show advanced options', 1;
RECONFIGURE;

-- Ad Hoc Distributed Queries özelliðini etkinleþtir
EXEC sp_configure 'Ad Hoc Distributed Queries', 1;
RECONFIGURE;

EXEC sp_configure 'Ad Hoc Distributed Queries';

SELECT * 
FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0', 
                'Excel 12.0;Database=C:\Kullanýcýlar\turka\Downloads\northwind(1).xlsx;HDR=YES', 
                'SELECT * FROM [Sheet1$]');


EXEC sp_configure 'Ad Hoc Distributed Queries', 0;
RECONFIGURE;

SELECT SERVERPROPERTY('Edition'), SERVERPROPERTY('ProductLevel'), SERVERPROPERTY('EngineEdition');


