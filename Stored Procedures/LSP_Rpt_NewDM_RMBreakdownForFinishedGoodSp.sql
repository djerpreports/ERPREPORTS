--CREATE PROCEDURE LSP_Rpt_NewDM_RMBreakdownForFinishedGoodSp (
DECLARE
	@StartDate				DateType = '01/01/2020'
  , @EndDate				DateType = '01/31/2020'
--) AS

BEGIN

	IF OBJECT_ID('tempdb..#itemPrice') IS NOT NULL
		DROP TABLE #itemPrice
	IF OBJECT_ID('tempdb..#BOMStdCost') IS NOT NULL
		DROP TABLE #BOMStdCost	
	IF OBJECT_ID('tempdb..#DMActualCost') IS NOT NULL
		DROP TABLE #DMActualCost
	
	CREATE TABLE #BOMStdCost (
		item				NVARCHAR(60)
	  , [Level]				INT
	  , Parent				NVARCHAR(20)
	  , oper_num			INT
	  , sequence			INT
	  , subsequence			NVARCHAR(50)
	  , matl				NVARCHAR(60)
	  , matl_qty			DECIMAL(18,10)
	  , matl_unit_cost		DECIMAL(18,10)
	  , pi_process_cost		DECIMAL(18,10)
	  , pi_resin_cost		DECIMAL(18,10)
	  , pi_hidden_profit	DECIMAL(18,10)
	  , sf_labr_cost		DECIMAL(18,10)
	  , sf_ovhd_cost		DECIMAL(18,10)
	  , fg_labr_cost		DECIMAL(18,10)
	  , fg_ovhd_cost		DECIMAL(18,10)
	)
	
	CREATE TABLE #DMActualCost (
		item						NVARCHAR(60)
	  , [Level]						INT
	  , Parent						NVARCHAR(20)
	  , oper_num					INT
	  , sequence					INT
	  , subsequence					NVARCHAR(50)
	  , matl						NVARCHAR(60)
	  , matl_qty					DECIMAL(18,8)
	  , lot_no						NVARCHAR(50)
	  , trans_date					DATETIME
	  , job_qty						BIGINT
	  , matl_unit_cost_usd			DECIMAL(18,8)
	  , matl_landed_cost_usd		DECIMAL(18,8)
	  , pi_fg_process_usd			DECIMAL(18,8)
	  , pi_resin_usd				DECIMAL(18,8)
	  , pi_vend_cost_usd			DECIMAL(18,8)
	  , pi_hidden_profit_usd		DECIMAL(18,8)
	  , sf_lbr_cost_usd				DECIMAL(18,8)
	  , sf_ovhd_cost_usd			DECIMAL(18,8)
	  , fg_lbr_cost_usd				DECIMAL(18,8)
	  , fg_ovhd_cost_usd			DECIMAL(18,8)
	  , matl_unit_cost_php			DECIMAL(18,8)
	  , matl_landed_cost_php		DECIMAL(18,8)
	  , pi_fg_process_php			DECIMAL(18,8)
	  , pi_resin_php				DECIMAL(18,8)
	  , pi_vend_cost_php			DECIMAL(18,8)
	  , pi_hidden_profit_php		DECIMAL(18,8)
	  , sf_lbr_cost_php				DECIMAL(18,8)
	  , sf_ovhd_cost_php			DECIMAL(18,8)
	  , fg_lbr_cost_php				DECIMAL(18,8)
	  , fg_ovhd_cost_php			DECIMAL(18,8)
	)
	
	DECLARE @ship_tran AS TABLE (
		TransDate			DateType
	  , Item				ItemType
	  , ItemDesc			DescriptionType
	  , ProductCode			ProductCodeType
	  , QtyShipped			QtyUnitType
	  , JobOrder			JobType
	  , JobSuffix			SuffixType
	  , PONumber			NVARCHAR(20)
	  , LotNumber			LotType
	  , FamilyCode			FamilyCodeType
	  , FamilyDesc			DescriptionType
	  , CONum				CoNumType
	  , COLine				CoLineType
	  , CustNum				CustNumType
	  , CustShipTo			CustSeqType
	  , CustomerName		NameType
	  , SalesUnitPrice		CostPrcType
	)
	
	DECLARE @shipped AS TABLE (
		TransDate			DateType
	  , Item				ItemType
	  , ItemDesc			DescriptionType
	  , ProductCode			ProductCodeType
	  , QtyShipped			QtyUnitType
	  , JobOrder			JobType
	  , JobSuffix			SuffixType
	  , PONumber			NVARCHAR(20)
	  , LotNumber			LotType
	  , FamilyCode			FamilyCodeType
	  , FamilyDesc			DescriptionType
	  , CONum				CoNumType
	  , COLine				CoLineType
	  , CustNum				CustNumType
	  , CustShipTo			CustSeqType
	  , CustomerName		NameType
	  , SalesUnitPrice		CostPrcType
	)

	DECLARE @FGReceipts AS TABLE (
		TransDate			DateType
	  , Item				ItemType
	  , ItemDesc			DescriptionType
	  , ProductCode			ProductCodeType
	  , QtyCompleted		QtyUnitType
	  , JobOrder			JobType
	  , JobSuffix			SuffixType
	  , PONumber			NVARCHAR(20)
	  , FamilyCode			FamilyCodeType
	  , FamilyDesc			DescriptionType
	  , CONum				CoNumType
	  , CustNum				CustNumType
	  , CustShipTo			CustSeqType
	  , CustomerName		NameType
	  , FGTransType			NVARCHAR(20)
	)
	
	DECLARE @FGReportSet AS TABLE (  
		TransDate				DateType
	  , PONum					NVARCHAR(20)
	  , JobOrder				JobType
	  , JobSuffix				SuffixType
	  , Item					ItemType
	  , ItemDesc				DescriptionType
	  , ProductCode				ProductCodeType
	  , Family					FamilyCodeType
	  , FamilyDesc				DescriptionType
	  , CONum					CoNumType
	  , CustNum					CustNumType
	  , ShipToCust				CustSeqType
	  , CustomerName			NameType
	  , FGTransType				NVARCHAR(25)
	  , QtyCompleted			QtyUnitType
	  
	  /****STANDARD COSTS****/
	  , StdMatlCost_PHP		AmountType
	  , StdLandedCost_PHP		AmountType
	  , StdResinCost_PHP		AmountType
	  , StdPIProcess_PHP		AmountType
	  , StdHiddenProfit_PHP		AmountType	  
	  , StdSFAdded_PHP			AmountType
	  , StdFGAdded_PHP			AmountType	  
	  , StdUnitCost_PHP			AmountType
	  
	  /****ACTUAL COSTS****/	  
	  , ActlMatlUnitCost_PHP	AmountType
	  , ActlLandedCost_PHP		AmountType
	  , ActlResinCost_PHP		AmountType
	  , ActlPIProcess_PHP		AmountType
	  , ActlHiddenProfit_PHP	AmountType	  
	  , ActlSFAdded_PHP			AmountType
	  , ActlFGAdded_PHP			AmountType	 
	  , ActlUnitCost_PHP		AmountType 
	)
	
	DECLARE
		@TransDate				DateType
	  , @PONum					NVARCHAR(20)
	  , @JobOrder				JobType
	  , @JobSuffix				SuffixType
	  , @Item					ItemType
	  , @ItemDesc				DescriptionType
	  , @QtyCompleted			QtyUnitType
	  , @ProductCode			ProductCodeType
	  , @Family					FamilyCodeType
	  , @FamilyDesc				DescriptionType
	  , @CONum					CoNumType
	  , @CustNum				CustNumType
	  , @ShipToCust				CustSeqType
	  , @CustomerName			NameType
	  , @FGTransType			NVARCHAR(25)
	  
	  , @EXWUnitCost			CostPrcType  
	  , @EXWCurrCode			CurrCodeType
	  , @ExchRate				ExchRateType  
	  
	  , @StdMatlCost			AmountType
	  , @StdLandedCost			AmountType
	  , @StdResinCost			AmountType  
	  , @StdPIProcess			AmountType
	  , @StdPIHiddenProfit		AmountType
	  , @StdSFAdded				AmountType
	  , @StdFGAdded				AmountType
	  
	  , @ActlMatlCostPHP		AmountType
	  , @ActlLandedCostPHP		AmountType
	  , @ActlResinCostPHP		AmountType  
	  , @ActlPIProcessPHP		AmountType
	  , @ActlPIHiddenProfitPHP	AmountType
	  , @ActlSFAddedPHP			AmountType
	  , @ActlFGAddedPHP			AmountType
	  
	SELECT @StartDate = ISNULL(@StartDate, GETDATE())
		 , @EndDate = ISNULL(@EndDate, GETDATE())

	INSERT INTO @ship_tran
	SELECT m.trans_date
	  , m.item
	  , i.description
	  , i.product_code
	  , m.qty
	  , (SELECT TOP(1) matltran2.ref_num
		  FROM matltran matltran2
		  WHERE m.lot = matltran2.lot AND matltran2.trans_type = 'F' AND m.item = matltran2.item
		  ORDER BY matltran2.trans_date DESC)
	  , (SELECT TOP(1) matltran2.ref_line_suf
	  FROM matltran matltran2
	  WHERE m.lot = matltran2.lot AND matltran2.trans_type = 'F' AND m.item = matltran2.item
	  ORDER BY matltran2.trans_date DESC)
	  , coi.Uf_ponum
	  , m.lot
	  , i.family_code
	  , f.description
	  , m.ref_num
	  , m.ref_line_suf
	  , c.cust_num
	  , c.cust_seq
	  , ca.name
	  , coi.price_conv
	FROM matltran AS m
		JOIN item AS i
			ON m.item = i.item
		JOIN famcode AS f
			ON i.family_code = f.family_code
		JOIN coitem AS coi
			ON m.ref_num = coi.co_num AND m.ref_line_suf = coi.co_line
		LEFT OUTER JOIN co AS c
			ON coi.co_num = c.co_num
		LEFT OUTER JOIN custaddr AS ca
			ON c.cust_num = ca.cust_num AND c.cust_seq = ca.cust_seq /*LEFT OUTER JOIN
	  matltran matltran2 ON m.lot = matltran2.lot AND matltran2.trans_type = 'F' AND m.item = matltran2.item*/
	WHERE m.trans_date BETWEEN dbo.MidnightOf(@StartDate) AND dbo.DayEndOf(@EndDate)
	  AND m.trans_type = 'S' AND m.ref_type = 'O'
	  AND m.item LIKE 'FG-%'
	  AND m.ref_num = '19-0002507';
	
	WITH CTE_ship AS
	(SELECT MAX(TransDate) AS TransDate
	  , Item
	  , ItemDesc
	  , ProductCode
	  , (SUM(QtyShipped) * (-1)) AS QtyShipped
	  , JobOrder
	  , JobSuffix
	  , PONumber
	  , LotNumber
	  , FamilyCode
	  , FamilyDesc
	  , CONum
	  , COLine
	  , CustNum
	  , CustShipTo
	  , CustomerName
	  , SalesUnitPrice
	FROM @ship_tran
	GROUP BY PONumber, CONum, COLine, Item, ItemDesc, ProductCode, JobOrder, JobSuffix, LotNumber
		, FamilyCode, FamilyDesc, CustNum, CustShipTo, CustomerName, SalesUnitPrice)
	
	INSERT INTO @FGReceipts
	SELECT m.trans_date
	  , m.item
	  , i.description
	  , i.product_code
	  , m.qty
	  , m.ref_num
	  , m.ref_line_suf
	  , coi.Uf_ponum
	  , i.family_code
	  , f.description
	  , coi.co_num
	  , c.cust_num
	  , c.cust_seq
	  , ca.name
	  , CASE WHEN coi.Uf_ponum LIKE '%RP%'
					OR coi.Uf_ponum LIKE '%R%'
					OR coi.Uf_ponum LIKE '%S%'
				THEN 'SAMPLE/REPAIR'
			 WHEN ship.TransDate IS NULL
				THEN 'STOCK ASSESSMENT'
			 ELSE 'FINISHED GOODS'
		END
	
	FROM matltran AS m
		JOIN item AS i
			ON m.item = i.item
		JOIN famcode AS f
			ON i.family_code = f.family_code
		JOIN job AS j
			ON m.ref_num = j.job AND m.ref_line_suf = j.suffix
		LEFT OUTER JOIN coitem AS coi
			ON j.ord_num = coi.co_num AND j.ord_line = coi.co_line
		LEFT OUTER JOIN co AS c
			ON coi.co_num = c.co_num
		LEFT OUTER JOIN custaddr AS ca
			ON c.cust_num = ca.cust_num AND c.cust_seq = ca.cust_seq
		LEFT OUTER JOIN CTE_ship AS ship
			ON coi.Uf_ponum = ship.PONumber AND m.ref_num = ship.JobOrder AND m.ref_line_suf = ship.JobSuffix
	
	WHERE m.trans_date BETWEEN dbo.MidnightOf(@StartDate) AND dbo.DayEndOf(@EndDate)
	  AND m.qty > 0 AND m.trans_type = 'F' AND m.ref_type = 'J'
	  AND m.item LIKE 'FG-%'
	  AND m.ref_num = '19-0002507';
	--  AND (m.ref_num LIKE '__-%' OR m.ref_num LIKE '__S-%' OR m.ref_num LIKE '__RP-%')
	  
	  
	SELECT TOP(1) WITH TIES
		   item
		 , effect_date
		 , curr_code
		 , unit_price1
	INTO #itemPrice
	FROM itemprice
	WHERE effect_date < @StartDate AND effect_date < @EndDate
	  AND item IN (SELECT item FROM @FGReceipts)
	ORDER BY ROW_NUMBER() OVER (PARTITION BY item ORDER BY effect_date DESC)
	
	DECLARE FGCrsr CURSOR FAST_FORWARD FOR
	SELECT MAX(TransDate) AS TransDate
	  , PONumber
	  , JobOrder
	  , JobSuffix
	  , Item
	  , ItemDesc
	  , ProductCode
	  , FamilyCode
	  , FamilyDesc
	  , MAX(CONum)
	  , MAX(CustNum)
	  , MAX(CustShipTo)
	  , MAX(CustomerName)
	  , MAX(FGTransType)
	  , SUM(QtyCompleted)
	  
	FROM @FGReceipts
	GROUP BY JobOrder, JobSuffix, PONumber, Item, ItemDesc, ProductCode, FamilyCode, FamilyDesc
	ORDER BY MAX(TransDate)


	OPEN FGCrsr
	FETCH FROM FGCrsr INTO
		@TransDate
	  , @PONum
	  , @JobOrder
	  , @JobSuffix
	  , @Item
	  , @ItemDesc
	  , @ProductCode
	  , @Family
	  , @FamilyDesc
	  , @CONum
	  , @CustNum
	  , @ShipToCust
	  , @CustomerName
	  , @FGTransType
	  , @QtyCompleted
	
	WHILE @@FETCH_STATUS = 0
	BEGIN
		
		TRUNCATE TABLE #BOMStdCost
	  
		INSERT INTO #BOMStdCost
		EXEC dbo.LSP_DM_StdCost_GetCurrentMatlCostingSp @Item, @TransDate
	
		SELECT @StdMatlCost	= ISNULL(matl_unit_cost, 0)
			 , @StdLandedCost = 0
			 , @StdResinCost = ISNULL(pi_resin_cost, 0)
			 , @StdPIProcess = ISNULL(pi_process_cost, 0)
			 , @StdPIHiddenProfit = ISNULL(pi_hidden_profit, 0)
			 , @StdSFAdded = ISNULL(sf_labr_cost, 0) + ISNULL(sf_ovhd_cost, 0)
			 , @StdFGAdded = ISNULL(fg_labr_cost, 0) + ISNULL(fg_ovhd_cost, 0)
		FROM #BOMStdCost
		WHERE [Level] = 0	
	
		TRUNCATE TABLE #DMActualCost
		
		INSERT INTO #DMActualCost
		EXEC dbo.LSP_DM_ActlCost_GetJobMatlTransCostingSp @JobOrder, @JobSuffix, @Item, @TransDate, @QtyCompleted
		
		SELECT @ActlMatlCostPHP	= ISNULL(matl_unit_cost_php,0) / ISNULL(job_qty,0)
			 , @ActlLandedCostPHP	= ISNULL(matl_landed_cost_php,0) / ISNULL(job_qty,0)
			 , @ActlResinCostPHP	= ISNULL(pi_resin_php,0) / ISNULL(job_qty,0)
			 , @ActlPIProcessPHP	= ISNULL(pi_fg_process_php,0) / ISNULL(job_qty,0)
			 , @ActlPIHiddenProfitPHP = ISNULL(pi_hidden_profit_php,0) / ISNULL(job_qty,0)
			 , @ActlSFAddedPHP		= (ISNULL(sf_lbr_cost_php,0) + ISNULL(sf_ovhd_cost_php,0)) / ISNULL(job_qty,0)
			 , @ActlFGAddedPHP		= (ISNULL(fg_lbr_cost_php,0) + ISNULL(fg_ovhd_cost_php,0)) / ISNULL(job_qty,0)
		FROM #DMActualCost
		WHERE [Level] = 0		
		
		SELECT @EXWUnitCost = ISNULL(unit_price1,0) / 1.2
			 , @EXWCurrCode = curr_code
		FROM #itemPrice
		WHERE item = @Item
		 
		IF @EXWCurrCode <> 'PHP'
		BEGIN		
			EXEC dbo.LSP_CurrencyConversionModSp @TransDate, @EXWCurrCode, 'PHP', @EXWUnitCost, @EXWUnitCost OUTPUT, @ExchRate OUTPUT
		END
		ELSE
		BEGIN
			EXEC dbo.LSP_ConvertUsdToPhpCurrencySp @TransDate, @ExchRate OUTPUT  
		END
		
		SELECT * FROM #DMActualCost
		
		SELECT @JobOrder
			 , ISNULL(@PONum, '')
			 , @QtyCompleted
			 , @Item
			 , @ItemDesc
			 , ac.matl
			 , ac.matl_qty
			 , ac.lot_no
			 , ac.matl_unit_cost_php
			 , ac.matl_landed_cost_php
			 , ac.pi_resin_php
			 , ac.pi_fg_process_php
			 , ac.pi_hidden_profit_php
			 , ISNULL(ac.sf_lbr_cost_php,0)
			 , ISNULL(ac.sf_ovhd_cost_php,0)
			 
			 , std.matl
			 , std.matl_qty
			 , std.matl_unit_cost * @ExchRate
			 , std.pi_resin_cost * @ExchRate
			 , std.pi_process_cost * @ExchRate
			 , std.pi_hidden_profit * @ExchRate
			 , ISNULL(std.sf_labr_cost,0)
			 , ISNULL(std.sf_ovhd_cost,0)
		FROM #DMActualCost AS ac
			LEFT OUTER JOIN #BOMStdCost AS std 
				ON ac.matl = std.matl 
					AND ac.item = std.item 
					AND ac.[Level] = std.[Level]
					AND ac.oper_num = std.oper_num
		WHERE ac.[Level] <> 0
		
		INSERT INTO @FGReportSet
		SELECT @TransDate
			 , ISNULL(@PONum, '')
			 , @JobOrder
			 , @JobSuffix
			 , @Item
			 , @ItemDesc
			 , @ProductCode
			 , @Family
			 , @FamilyDesc
			 , ISNULL(@CONum, '')
			 , ISNULL(@CustNum, '')
			 , ISNULL(@ShipToCust, '')
			 , ISNULL(@CustomerName, '')
			 , @FGTransType
			 , @QtyCompleted
			 , @StdMatlCost * ISNULL(@ExchRate, 0)
			 , @StdLandedCost * ISNULL(@ExchRate, 0)
			 , @StdResinCost * ISNULL(@ExchRate, 0)
			 , @StdPIProcess * ISNULL(@ExchRate, 0)
			 , @StdPIHiddenProfit * ISNULL(@ExchRate, 0)
			 , @StdSFAdded
			 , @StdFGAdded
			 , ((@StdMatlCost + @StdLandedCost 
				+ @StdResinCost + @StdPIProcess + @StdPIHiddenProfit) * ISNULL(@ExchRate, 0))
				+ @StdSFAdded + @StdFGAdded
			 , @ActlMatlCostPHP
			 , @ActlLandedCostPHP
			 , @ActlResinCostPHP
			 , @ActlPIProcessPHP
			 , @ActlPIHiddenProfitPHP
			 , @ActlSFAddedPHP
			 , @ActlFGAddedPHP
			 , (@ActlMatlCostPHP + @ActlLandedCostPHP 
				+ @ActlResinCostPHP + @ActlPIProcessPHP + @ActlPIHiddenProfitPHP 
				+ @ActlSFAddedPHP + @ActlFGAddedPHP)
	
		SELECT @EXWUnitCost = 0
			 , @EXWCurrCode = ''
			 , @ExchRate = 0
			 , @StdMatlCost = 0
			 , @StdLandedCost = 0
			 , @StdResinCost = 0
			 , @StdPIProcess = 0
			 , @StdPIHiddenProfit = 0
			 , @StdSFAdded = 0
			 , @StdFGAdded = 0
			 , @ActlMatlCostPHP = 0
			 , @ActlLandedCostPHP = 0
			 , @ActlResinCostPHP = 0
			 , @ActlPIProcessPHP = 0
			 , @ActlPIHiddenProfitPHP = 0
			 , @ActlSFAddedPHP = 0
			 , @ActlFGAddedPHP = 0
	
		FETCH NEXT FROM FGCrsr INTO
			@TransDate
		  , @PONum
		  , @JobOrder
		  , @JobSuffix
		  , @Item
		  , @ItemDesc
		  , @ProductCode
		  , @Family
		  , @FamilyDesc
		  , @CONum
		  , @CustNum
		  , @ShipToCust
		  , @CustomerName
		  , @FGTransType
		  , @QtyCompleted		  
	
	END
	
	CLOSE FGCrsr
	DEALLOCATE FGCrsr	

	SELECT * 
	FROM @FGReportSet

END