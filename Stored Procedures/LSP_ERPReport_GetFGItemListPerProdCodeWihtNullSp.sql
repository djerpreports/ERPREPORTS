--EXEC dbo.LSP_ERPReport_GetFGItemListPerProdCodeWihtNullSp '',''

ALTER PROCEDURE LSP_ERPReport_GetFGItemListPerProdCodeWihtNullSp (  
--DECLARE  
	@StartProdCode		ProductCodeType = NULL  
  , @EndProdCode		ProductCodeType = NULL  
) AS  
  
BEGIN
  
	SELECT @StartProdCode = ISNULL(NULLIF(@StartProdCode,''), dbo.LowCharacter())  
		 , @EndProdCode = ISNULL(NULLIF(@EndProdCode,''), dbo.HighCharacter())  
	  
	SELECT item
		 , [description]
	FROM item  
	WHERE (product_code BETWEEN @StartProdCode AND @EndProdCode)
	  AND item LIKE 'FG-%'  
	UNION  
	SELECT '' AS item
		 , '' AS [description]
	ORDER BY item  
  
END