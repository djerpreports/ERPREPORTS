--CREATE PROCEDURE LSP_Rpt_NewDM_SalesAndSampleJOReportSp (
DECLARE
	@StartDate					DateType	= '12/01/2019'
  , @EndDate					DateType	= '01/31/2020'
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
	
	DECLARE @ShipTrans AS TABLE (  
		TransDate				DateType
	  , Item					ItemType
	  , ItemDesc				DescriptionType
	  , ProductCode				ProductCodeType
	  , QtyShipped				QtyUnitType
	  , JobOrder				JobType
	  , JobSuffix				SuffixType
	  , PONumber				NVARCHAR(20)
	  , LotNumber				LotType
	  , FamilyCode				FamilyCodeType
	  , FamilyDesc				DescriptionType
	  , CONum					CoNumType
	  , COLine					CoLineType
	  , CustNum					CustNumType
	  , CustShipTo				CustSeqType
	  , CustomerName			NameType
	  , SalesUnitPrice			CostPrcType
	   , TransNum				MatlTransNumType
	)
	
	DECLARE @SalesReport AS TABLE (  
		TransDate				DateType  
	  , Item					ItemType  
	  , ItemDesc				DescriptionType  
	  , ProductCode				ProductCodeType  
	  , Family					FamilyCodeType  
	  , FamilyDesc				DescriptionType  
	  , PONum					NVARCHAR(20)
	  , LotNo					LotType
	  , JobOrder				JobType  
	  , JobSuffix				SuffixType  
	  , CONum					CoNumType
	  , COLine					CoLineType  
	  , CustNum					CustNumType
	  , ShipToCust				CustSeqType  
	  , CustomerName			NameType  
	  , QtyShipped				QtyUnitType  
	  
	  , SalesPrice				AmountType  
	  , SalesPriceConv			AmountType  	  
	  
	  /****STANDARD COSTS****/
	  , StdMatlCost_PHP			AmountType
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

	  , ShipCategory			NVARCHAR(10)  
	  , Recoverable				INT
	  , JobRemarks				NVARCHAR(200)
	)
	
	DECLARE
		@TransDate				DateType
	  , @Item					ItemType
	  , @ItemDesc				DescriptionType
	  , @ProductCode			ProductCodeType
	  , @Family					FamilyCodeType
	  , @FamilyDesc				DescriptionType
	  , @PONum					NVARCHAR(20)
	  , @JobOrder				JobType
	  , @JobSuffix				SuffixType
	  , @LotNo					LotType
	  , @CONum					CoNumType
	  , @COLine					CoLineType  
	  , @CustNum				CustNumType
	  , @ShipToCust				CustSeqType
	  , @CustomerName			NameType
	  , @QtyCompleted			QtyUnitType
	  , @SalesUnitPrice			CostPrcType
	  
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
	  , @Recoverable			INT
	  , @JobRemarks				NVARCHAR(200)
	  , @IsJobExists			BIT
	  
	  , @StdSFLbrCst				DECIMAL(18,10)
	  , @StdSFOvhdCst				DECIMAL(18,10)
	  , @StdFGLbrCst				DECIMAL(18,10)
	  , @StdFGOvhdCst				DECIMAL(18,10)
	  , @ActlSFLbrCst				DECIMAL(18,10)
	  , @ActlSFOvhdCst				DECIMAL(18,10)
	  , @ActlFGLbrCst				DECIMAL(18,10)
	  , @ActlFGOvhdCst				DECIMAL(18,10)
	  , @PIVendCost				DECIMAL(18,10)
	
	SELECT @StartDate = ISNULL(@StartDate, GETDATE())
		 , @EndDate = ISNULL(@EndDate, GETDATE())

	INSERT INTO @ShipTrans
	SELECT m.trans_date
	  , m.item
	  , i.description
	  , i.product_code
	  , m.qty
	  , (SELECT TOP(1) m2.ref_num
		 FROM matltran AS m2
		 WHERE m.lot = m2.lot AND m2.trans_type = 'F' AND m.item = m2.item
		 ORDER BY m2.trans_date DESC)
	  , (SELECT TOP(1) m2.ref_line_suf
		 FROM matltran AS m2
		 WHERE m.lot = m2.lot AND m2.trans_type = 'F' AND m.item = m2.item
		 ORDER BY m2.trans_date DESC)
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
	, m.trans_num
	FROM matltran AS m
		JOIN item AS i
			ON m.item = i.item
		JOIN famcode AS f
			ON i.family_code = f.family_code
		JOIN coitem AS coi
			ON m.ref_num = coi.co_num AND m.ref_line_suf = coi.co_line
		JOIN co AS c
			ON coi.co_num = c.co_num
		JOIN custaddr AS ca
			ON c.cust_num = ca.cust_num AND c.cust_seq = ca.cust_seq
		JOIN do_seq AS dos
			ON coi.co_num = dos.ref_num AND coi.co_line  = dos.ref_line
		JOIN do_hdr AS doh
			ON dos.do_num = doh.do_num
	WHERE m.trans_date BETWEEN dbo.MidnightOf(@StartDate) AND dbo.DayEndOf(@EndDate)
	  AND m.trans_type = 'S' AND m.ref_type = 'O' AND doh.stat = 'A'
	  
	GROUP BY m.trans_date, m.item, i.description, i.product_code, m.qty, coi.Uf_ponum, m.lot, i.family_code, f.description
	  , m.ref_num, m.ref_line_suf, c.cust_num, c.cust_seq, ca.name, coi.price_conv, m.trans_num
	
	UNION
	
	SELECT m.trans_date
		 , m.item
		 , i.description
		 , i.product_code
		 , m.qty
		 , (SELECT TOP(1) m2.ref_num
			FROM matltran AS m2
			WHERE m.lot = m2.lot AND m2.trans_type = 'F' AND m.item = m2.item
			ORDER BY m2.trans_date DESC)
		 , (SELECT TOP(1) m2.ref_line_suf
			FROM matltran AS m2
			WHERE m.lot = m2.lot AND m2.trans_type = 'F' AND m.item = m2.item
			ORDER BY m2.trans_date DESC)
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
		 , m.trans_num
	FROM matltran AS m
		JOIN item AS i
			ON m.item = i.item
		JOIN famcode AS f
			ON i.family_code = f.family_code
		LEFT OUTER JOIN coitem AS coi
			ON m.ref_num = coi.co_num AND m.ref_line_suf = coi.co_line
		JOIN co AS c
			ON coi.co_num = c.co_num
		JOIN custaddr AS ca
			ON c.cust_num = ca.cust_num AND c.cust_seq = ca.cust_seq
	
	WHERE m.trans_date BETWEEN dbo.MidnightOf(@StartDate) AND dbo.DayEndOf(@EndDate)
	  AND m.trans_type = 'S' AND m.ref_type = 'O' AND ca.name LIKE 'LSP%SECTION'
	GROUP BY m.trans_date, m.item, i.description, i.product_code, m.qty, coi.Uf_ponum, m.lot, i.family_code, f.description
	  , m.ref_num, m.ref_line_suf, c.cust_num, c.cust_seq, ca.name, coi.price_conv, m.trans_num
	
	SELECT TOP(1) WITH TIES
		   item
		 , effect_date
		 , curr_code
		 , unit_price1
	INTO #itemPrice
	FROM itemprice
	WHERE effect_date < @StartDate AND effect_date < @EndDate
	  AND item IN (SELECT item FROM @ShipTrans)
	ORDER BY ROW_NUMBER() OVER (PARTITION BY item ORDER BY effect_date DESC)
	
	DECLARE shipCrsr CURSOR FAST_FORWARD FOR
	SELECT MAX(TransDate) TransDate
	  , Item
	  , ItemDesc
	  , ProductCode
	  , FamilyCode
	  , FamilyDesc
	  , PONumber
	  , LotNumber
	  , JobOrder
	  , JobSuffix
	  , CONum
	  , COLine
	  , CustNum
	  , CustShipTo
	  , CustomerName
	  , SalesUnitPrice
	  , (SUM(QtyShipped) * (-1)) QtyShipped
	
	FROM @ShipTrans
	GROUP BY PONumber, CONum, COLine, Item, ItemDesc, ProductCode, JobOrder, JobSuffix, LotNumber
		, FamilyCode, FamilyDesc, CustNum, CustShipTo, CustomerName, SalesUnitPrice
	ORDER BY PONumber

	OPEN shipCrsr
	FETCH FROM shipCrsr INTO
		@TransDate
	  , @Item
	  , @ItemDesc
	  , @ProductCode
	  , @Family
	  , @FamilyDesc
	  , @PONum
	  , @LotNo
	  , @JobOrder
	  , @JobSuffix
	  , @CONum
	  , @COLine
	  , @CustNum
	  , @ShipToCust
	  , @CustomerName	  
	  , @SalesUnitPrice
	  , @QtyCompleted
	  
	WHILE @@FETCH_STATUS = 0
	BEGIN
	
		SELECT @Recoverable = ISNULL(Uf_is_recoverable, 0)
			 , @JobRemarks = ISNULL(Uf_job_remarks, '')
			 , @IsJobExists = CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END
		FROM job  
		WHERE job = @JobOrder 
		  AND suffix = @JobSuffix
		  AND item = @Item
		GROUP BY job, suffix, item, Uf_is_recoverable, Uf_job_remarks
	
		TRUNCATE TABLE #BOMStdCost
		TRUNCATE TABLE #DMActualCost
	
		IF @IsJobExists = 1
		BEGIN
			INSERT INTO #BOMStdCost
			EXEC dbo.LSP_DM_StdCost_GetCurrentMatlCostingSp @Item, @TransDate
		
			SELECT @StdMatlCost	= matl_unit_cost
				 , @StdLandedCost = 0
				 , @StdResinCost = pi_resin_cost
				 , @StdPIProcess = pi_process_cost
				 , @StdPIHiddenProfit = pi_hidden_profit
				 , @StdSFAdded = sf_labr_cost + sf_ovhd_cost
				 , @StdFGAdded = fg_labr_cost + fg_ovhd_cost
			FROM #BOMStdCost
			WHERE [Level] = 0
		
			INSERT INTO #DMActualCost
			EXEC dbo.LSP_DM_ActlCost_GetJobMatlTransCostingSp @JobOrder, @JobSuffix, @Item, @TransDate, @QtyCompleted
			
			SELECT @ActlMatlCostPHP	= matl_unit_cost_php / job_qty
				 , @ActlLandedCostPHP	= matl_landed_cost_php / job_qty
				 , @ActlResinCostPHP	= pi_resin_php / job_qty
				 , @ActlPIProcessPHP	= pi_fg_process_php / job_qty
				 , @ActlPIHiddenProfitPHP = pi_hidden_profit_php / job_qty
				 , @ActlSFAddedPHP		= (sf_lbr_cost_php + sf_ovhd_cost_php) / job_qty
				 , @ActlFGAddedPHP		= (fg_lbr_cost_php + fg_ovhd_cost_php) / job_qty
			FROM #DMActualCost
			WHERE [Level] = 0
		END
		ELSE
		BEGIN
			EXEC dbo.LSP_StdCost_GetMatlCostingSp @Item, @TransDate
					, @StdMatlCost OUTPUT, @StdPIProcess OUTPUT, @StdResinCost OUTPUT, @StdPIHiddenProfit OUTPUT
					, @StdSFLbrCst OUTPUT, @StdSFOvhdCst OUTPUT, @StdFGLbrCst OUTPUT, @StdFGOvhdCst OUTPUT
			
			EXEC dbo.LSP_ActlCost_GetMatlCostingSp @Item, @LotNo, @TransDate
					  , 0 
					  , 0 , 0 
					  , 0 , 0 , 0 , 0
					  , 0 , 0
					  , 0 , 0
					  , @ActlMatlCostPHP OUTPUT, @ActlLandedCostPHP OUTPUT
					  , @ActlPIProcessPHP OUTPUT, @ActlResinCostPHP OUTPUT, @PIVendCost OUTPUT, @ActlPIHiddenProfitPHP OUTPUT
					  , @ActlSFLbrCst OUTPUT, @ActlSFOvhdCst OUTPUT
					  , @ActlFGLbrCst OUTPUT, @ActlFGOvhdCst OUTPUT
			
			SELECT @StdSFAdded = @StdSFLbrCst + @StdSFOvhdCst
				 , @StdFGAdded = @StdFGLbrCst + @StdFGOvhdCst
				 , @ActlSFAddedPHP = @ActlSFLbrCst + @ActlSFOvhdCst
				 , @ActlFGAddedPHP = @ActlFGLbrCst + @ActlFGOvhdCst
		
		END
		
		SELECT @EXWUnitCost = unit_price1 / 1.2
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
		
		
		INSERT INTO @SalesReport
		SELECT @TransDate
			 , @Item
			 , @ItemDesc
			 , @ProductCode
			 , @Family
			 , @FamilyDesc
			 , @PONum
			 , @LotNo
			 , ISNULL(@JobOrder, '')
			 , ISNULL(@JobSuffix, '')
			 , @CONum
			 , @COLine
			 , @CustNum
			 , @ShipToCust
			 , @CustomerName	  
			 , @QtyCompleted
			 , @SalesUnitPrice
			 , @SalesUnitPrice * ISNULL(@ExchRate, 0)
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
			 , CASE WHEN @CustomerName LIKE 'LSP%SECTION'
						THEN 'Sample JO'
					ELSE 'Sales' END
		     , @Recoverable
		     , @JobRemarks
		   
		   
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
			 , @Recoverable = ''
		     , @JobRemarks = ''
		     , @IsJobExists = 0
			 , @StdSFLbrCst = 0
			 , @StdSFOvhdCst = 0
			 , @StdFGLbrCst = 0
			 , @StdFGOvhdCst = 0
			 , @ActlSFLbrCst = 0
			 , @ActlSFOvhdCst = 0
			 , @ActlFGLbrCst = 0
			 , @ActlFGOvhdCst = 0
			 , @PIVendCost = 0
	
		FETCH NEXT FROM shipCrsr INTO
			@TransDate
		  , @Item
		  , @ItemDesc
		  , @ProductCode
		  , @Family
		  , @FamilyDesc
		  , @PONum
		  , @LotNo
		  , @JobOrder
		  , @JobSuffix
		  , @CONum
		  , @COLine
		  , @CustNum
		  , @ShipToCust
		  , @CustomerName	  
		  , @SalesUnitPrice
		  , @QtyCompleted
	
	
	END
	
	CLOSE shipCrsr
	DEALLOCATE shipCrsr


	--INSERT INTO DROP TABLE [Rpt_SalesSampleJO]
	SELECT * 	
	--INTO [Rpt_SalesSampleJO]
	FROM @SalesReport

END