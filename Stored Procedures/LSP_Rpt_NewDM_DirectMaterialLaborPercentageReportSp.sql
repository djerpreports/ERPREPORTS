--EXEC dbo.LSP_Rpt_NewDM_DirectMaterialLaborPercentageReportSp '05/01/2020', '05/31/2020', '', '', '',''

ALTER PROCEDURE LSP_Rpt_NewDM_DirectMaterialLaborPercentageReportSp (  
--DECLARE
	@StartDate			DateType		--= '05/01/2020'
  , @EndDate			DateType		--= '05/31/2020'
  , @StartProdCode		ProductCodeType --= 'FG-RS204'
  , @EndProdCode		ProductCodeType --= 'FG-RS204'
  , @StartModel			ItemType		--= ''
  , @EndModel			ItemType		--= ''
) AS
  
BEGIN TRANSACTION
	IF OBJECT_ID('tempdb..#itemPrice') IS NOT NULL
		DROP TABLE #itemPrice
	
	DECLARE  
		@TransDate		DateType
	  , @Item			ItemType
	  , @ItemDesc		DescriptionType
	  , @ProdCode		ProductCodeType  
	  , @QtyCompleted   QtyUnitType  
	  , @JobOrder		JobType  
	  , @JobSuffix		SuffixType  
	  , @FamilyCode		FamilyCodeType  
	  , @FamilyDesc		DescriptionType  
	  --, @EXWCostConv	AmountType  
	  , @EXWUnitCost		CostPrcType  
	  , @EXWCurrCode		CurrCodeType
	  , @ExchRate			ExchRateType  
	  , @StdMatlCost		AmountType
	  , @StdLandedCost		AmountType
	  , @StdResinCost		AmountType  
	  , @StdPIProcess		AmountType
	  , @StdPIHiddenProfit	AmountType
	  , @StdSFLbr			AmountType
	  , @StdSFOvhd			AmountType
	  , @StdFGLbr			AmountType
	  , @StdFGOvhd			AmountType
	  
	  , @ActlMatlCost		AmountType
	  , @ActlLandedCost		AmountType
	  , @ActlResinCost		AmountType  
	  , @ActlPIProcess		AmountType
	  , @ActlPIHiddenProfit	AmountType
	  , @ActlSFLbr			AmountType
	  , @ActlSFOvhd			AmountType
	  , @ActlFGLbr			AmountType
	  , @ActlFGOvhd			AmountType
	  	  
	DECLARE @report_set AS TABLE (  
		trans_date			DateType  
	  , item				ItemType  
	  , description			DescriptionType  
	  , product_code		ProductCodeType  
	  , fam_code			FamilyCodeType  
	  , fam_desc			DescriptionType  
	  , qty_completed		QtyUnitType  
	  , exw_unit			CostPrcType  
	  , produced_amt		AmountType  
	  , std_rm_cost			CostPrcType  
	  , std_lbr_cost		CostPrcType  
	  , standard_cost		AmountType  
	  , actual_cost			AmountType  
	  , actl_rm_cost		CostPrcType  
	  , actl_lbr_cost		AmountType  
	)  
	  
	SELECT @StartDate = dbo.MidnightOf(ISNULL(@StartDate, GETDATE()))
		 , @EndDate = dbo.DayEndOf(ISNULL(@EndDate, GETDATE()))
		 , @StartProdCode = ISNULL(NULLIF(@StartProdCode,''), (SELECT TOP(1) product_code FROM prodcode WHERE product_code LIKE 'FG-%' ORDER BY product_code ASC))
		 , @EndProdCode = ISNULL(NULLIF(@EndProdCode,''), (SELECT TOP(1) product_code FROM prodcode WHERE product_code LIKE 'FG-%' ORDER BY product_code DESC))
		 
	SELECT @StartModel = ISNULL(NULLIF(@StartModel,''), (SELECT TOP(1) item FROM item WHERE product_code BETWEEN @StartProdCode AND @EndProdCode ORDER BY item ASC))
		 , @EndModel = ISNULL(NULLIF(@EndModel,''), (SELECT TOP(1) item FROM item WHERE product_code BETWEEN @StartProdCode AND @EndProdCode ORDER BY item DESC))
	
	
	EXEC dbo.LSP_NewDM_GetFilteredFinishedGoodsTransactionSp @StartDate, @EndDate, @StartProdCode, @EndProdCode, @StartModel, @EndModel  
	
	SELECT TOP(1) WITH TIES
		   item
		 , effect_date
		 , curr_code
		 , unit_price1
	INTO #itemPrice
	FROM itemprice
	WHERE effect_date < @StartDate AND effect_date < @EndDate
	  AND item IN (SELECT item FROM ##FGReceipts)
	ORDER BY ROW_NUMBER() OVER (PARTITION BY item ORDER BY effect_date DESC)
	
	DECLARE dmlbrCrsr CURSOR FAST_FORWARD FOR  
	SELECT *  
	FROM ##FGReceipts
	
	OPEN dmlbrCrsr  
	FETCH FROM dmlbrCrsr INTO  
		@TransDate
	  , @Item
	  , @ItemDesc
	  , @ProdCode
	  , @QtyCompleted
	  , @JobOrder
	  , @JobSuffix
	  , @FamilyCode
	  , @FamilyDesc
	  
	WHILE (@@FETCH_STATUS = 0)  
	BEGIN  
	  
		EXEC dbo.LSP_DM_StdCost_GetCurrentMatlCostingSp @Item, @TransDate
		
		SELECT @StdMatlCost	= matl_unit_cost * @QtyCompleted
			 , @StdLandedCost = 0
			 , @StdResinCost = pi_resin_cost * @QtyCompleted
			 , @StdPIProcess = pi_process_cost * @QtyCompleted
			 , @StdPIHiddenProfit = pi_hidden_profit * @QtyCompleted
			 , @StdSFLbr = sf_labr_cost * @QtyCompleted
			 , @StdSFOvhd = sf_ovhd_cost * @QtyCompleted
			 , @StdFGLbr = fg_labr_cost * @QtyCompleted
			 , @StdFGOvhd = fg_ovhd_cost * @QtyCompleted
		FROM ##BOMCost
		WHERE [Level] = 0
		
		EXEC dbo.LSP_DM_ActlCost_GetJobMatlTransCostingSp @JobOrder, @JobSuffix, @Item, @TransDate, @QtyCompleted
	  
		SELECT @ActlMatlCost	= matl_unit_cost_php
			 , @ActlLandedCost	= matl_landed_cost_php
			 , @ActlResinCost	= pi_resin_php
			 , @ActlPIProcess	= pi_fg_process_php
			 , @ActlPIHiddenProfit = pi_hidden_profit_php
			 , @ActlSFLbr		= sf_lbr_cost_php
			 , @ActlSFOvhd		= sf_ovhd_cost_php
			 , @ActlFGLbr		= fg_lbr_cost_php
			 , @ActlFGOvhd		= fg_ovhd_cost_php
		FROM ##ActualCost
		WHERE [Level] = 0	  
		 
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
		
		INSERT INTO @report_set (  
			 trans_date  
		   , item  
		   , [description]
		   , product_code  
		   , fam_code  
		   , fam_desc  
		   , qty_completed  
		   , exw_unit  
		   , produced_amt  
		   , std_rm_cost  
		   , std_lbr_cost  
		   , standard_cost  
		   , actual_cost  
		   , actl_rm_cost  
		   , actl_lbr_cost  
		  )  
		SELECT @TransDate  
			 , @Item  
			 , @ItemDesc  
			 , @ProdCode  
			 , @FamilyCode  
			 , @FamilyDesc  
			 , @QtyCompleted  
			 , @EXWUnitCost
			 , @EXWUnitCost * @QtyCompleted
			 , (@StdMatlCost + @StdLandedCost + @StdResinCost + @StdPIProcess  + @StdPIHiddenProfit) * @ExchRate
			 , (@StdSFLbr + @StdSFOvhd + @StdFGLbr + @StdFGOvhd)  --* @ExchRate
			 , ((@StdMatlCost + @StdResinCost + @StdPIProcess  + @StdLandedCost + @StdPIHiddenProfit) * @ExchRate) 
					+ (@StdSFLbr + @StdSFOvhd + @StdFGLbr + @StdFGOvhd) --standard cost  
			 , @ActlMatlCost + @ActlLandedCost + @ActlResinCost + @ActlPIProcess  + @ActlPIHiddenProfit + @ActlSFLbr + @ActlSFOvhd + @ActlFGLbr + @ActlFGOvhd --actual cost  
			 , @ActlMatlCost + @ActlLandedCost + @ActlResinCost + @ActlPIProcess  + @ActlPIHiddenProfit--actual RM cost  
			 , @ActlSFLbr + @ActlSFOvhd + @ActlFGLbr + @ActlFGOvhd
	 
		SELECT @StdMatlCost			= 0
			 , @StdLandedCost 		= 0
			 , @StdResinCost 		= 0
			 , @StdPIProcess 		= 0
			 , @StdPIHiddenProfit 	= 0
			 , @StdSFLbr 			= 0
			 , @StdSFOvhd 			= 0
			 , @StdFGLbr 			= 0
			 , @StdFGOvhd 			= 0
			 , @ActlMatlCost		= 0
			 , @ActlLandedCost		= 0
			 , @ActlResinCost		= 0
			 , @ActlPIProcess		= 0
			 , @ActlPIHiddenProfit 	= 0
			 , @ActlSFLbr			= 0
			 , @ActlSFOvhd			= 0
			 , @ActlFGLbr			= 0
			 , @ActlFGOvhd			= 0
	 
		FETCH NEXT FROM dmlbrCrsr INTO  
			 @TransDate  
		   , @Item  
		   , @ItemDesc  
		   , @ProdCode  
		   , @QtyCompleted  
		   , @JobOrder  
		   , @JobSuffix  
		   , @FamilyCode  
		   , @FamilyDesc  
	  
	END  
	  
	CLOSE dmlbrCrsr  
	DEALLOCATE dmlbrCrsr  
	  
	SELECT * FROM @report_set  
	--SELECT * FROM @FinishedTrans  
  
COMMIT TRANSACTION