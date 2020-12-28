ALTER PROCEDURE LSP_NewDM_GetFilteredFinishedGoodsTransactionSp (
--DECLARE
	@StartDate			DateType		--= '01/18/2020'
  , @EndDate			DateType		--= '01/19/2020'
  , @StartProdCode		ProductCodeType --= 'FG-DK100'
  , @EndProdCode		ProductCodeType --= 'FG-DK100'
  , @StartModel			ItemType		--= ''
  , @EndModel			ItemType		--= ''
) AS

BEGIN

	IF OBJECT_ID('tempdb..##FGReceipts') IS NOT NULL
		DROP TABLE ##FGReceipts

	SELECT @StartDate = ISNULL(@StartDate, GETDATE())
		 , @EndDate = ISNULL(@EndDate, GETDATE())
		 , @StartProdCode = ISNULL(NULLIF(@StartProdCode,''), (SELECT TOP(1) product_code FROM prodcode WHERE product_code LIKE 'FG-%' ORDER BY product_code ASC))
		 , @EndProdCode = ISNULL(NULLIF(@EndProdCode,''), (SELECT TOP(1) product_code FROM prodcode WHERE product_code LIKE 'FG-%' ORDER BY product_code DESC))
		 
	SELECT @StartModel = ISNULL(NULLIF(@StartModel,''), (SELECT TOP(1) item FROM item WHERE product_code BETWEEN @StartProdCode AND @EndProdCode ORDER BY item ASC))
		 , @EndModel = ISNULL(NULLIF(@EndModel,''), (SELECT TOP(1) item FROM item WHERE product_code BETWEEN @StartProdCode AND @EndProdCode ORDER BY item DESC))
	
	SELECT MAX(m.trans_date) AS TransDate
		 , m.item
		 , i.description AS ItemDesc
		 , i.product_code AS ProductCode
		 , SUM(m.qty) AS QtyCompleted
		 , m.ref_num AS JobOrder
		 , m.ref_line_suf AS JobSuffix
		 , i.family_code AS FamilyCode
		 , f.description AS FamilyDesc
	INTO ##FGReceipts
	FROM matltran AS m
		JOIN item AS i
			ON m.item = i.item 
		JOIN famcode AS f
			ON i.family_code = f.family_code 

	WHERE m.trans_date BETWEEN dbo.MidnightOf(@StartDate) AND dbo.DayEndOf(@EndDate)
	  AND m.qty > 0 AND m.trans_type = 'F' AND m.ref_type = 'J'
	  AND m.item LIKE 'FG-%' 
	--  AND (m.ref_num LIKE '__-%' OR m.ref_num LIKE '__S-%' OR m.ref_num LIKE '__RP-%')
	  AND i.product_code BETWEEN @StartProdCode AND @EndProdCode
	  AND m.item BETWEEN @StartModel AND @EndModel
	GROUP BY m.item, i.[description], i.product_code, m.ref_num, m.ref_line_suf, i.family_code, f.[description]

	--SELECT TransDate
	--	 , Item
	--	 , ItemDesc
	--	 , ProductCode
	--	 , QtyCompleted
	--	 , JobOrder
	--	 , JobSuffix
	--	 , FamilyCode
	--	 , FamilyDesc
	--FROM ##FGReceipts
	

END