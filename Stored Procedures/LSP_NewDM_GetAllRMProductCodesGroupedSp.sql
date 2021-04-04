CREATE PROCEDURE LSP_NewDM_GetAllRMProductCodesGroupedSp  
  
AS  
  
BEGIN 
  
	SELECT REPLACE(REPLACE(product_code, 'RM-', ''), 'SA-', '') AS ProductCode  
	  , REPLACE(REPLACE(product_code, 'RM-', ''), 'SA-', '') AS [Description]
	FROM            prodcode  
	WHERE        product_code LIKE 'RM-%' OR  
							 product_code LIKE 'SA-%' OR  
							 product_code = 'FG-PI' OR  
							 product_code = 'PS-RM'  
	GROUP BY REPLACE(REPLACE(product_code, 'RM-', ''), 'SA-', '')  
	UNION  
	SELECT 'ALL', 'ALL'
	FROM prodcode  
	  
	ORDER BY Description  
  
END