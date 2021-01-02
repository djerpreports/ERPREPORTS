--EXEC dbo.LSP_ERPReport_GetFGItemListPerProdCodeWihtNullSp '','', ''

ALTER PROCEDURE LSP_ERPReport_GetFGItemListPerProdCodeWihtNullSp (  
--DECLARE  
	@StartProdCode		ProductCodeType --= ''
  , @EndProdCode		ProductCodeType --= ''
  , @Search				NVARCHAR(100)	--= 'RS'
) AS  
  
BEGIN
  
	SELECT @StartProdCode = ISNULL(NULLIF(@StartProdCode,''), dbo.LowCharacter())
		 , @EndProdCode = ISNULL(NULLIF(@EndProdCode,''), dbo.HighCharacter())
		 , @Search = ISNULL(NULLIF(@Search,''), '')
	  
	SELECT item
		 , [description]
	FROM item  
	WHERE (product_code BETWEEN @StartProdCode AND @EndProdCode)
	  AND item LIKE 'FG-%'
	  AND product_code LIKE '%'+@Search+'%'
	UNION  
	SELECT '' AS item
		 , '' AS [description]
	ORDER BY item  
  
END
