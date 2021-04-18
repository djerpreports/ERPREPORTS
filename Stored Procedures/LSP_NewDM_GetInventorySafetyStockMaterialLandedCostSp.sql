ALTER PROCEDURE LSP_NewDM_GetInventorySafetyStockMaterialLandedCostSp (  
--DECLARE  
	@ProdCode				ProductCodeType --= 'DK2300'
  , @InvtyMaterialCost		AmountType		OUTPUT  
  , @InvtyLandedCost		AmountType		OUTPUT  
  , @SafetyMaterialCost		AmountType		OUTPUT  
--  , @SafetyLandedCost  AmountType  OUTPUT  
) AS  
  
BEGIN

	IF OBJECT_ID('tempdb..#DMActualCost') IS NOT NULL
		DROP TABLE #DMActualCost
  
	DECLARE   
		@RMProdCode				ProductCodeType  
	  , @SFProdCode				ProductCodeType  
	  , @Item					ItemType  
	  , @ItemProdCode			ProductCodeType  
	  , @SafetyStockQty			QtyUnitType  
	  , @LotNumber				LotType 
	  , @LotCreateDate			DateType
	  , @LotQty					QtyUnitType  
	  , @LotMatlCost			AmountType  
	  , @LotLandedCost			AmountType  
	  , @StdItemPrice			DECIMAL(18,8)
	  , @StdItemPrice_PHP		DECIMAL(18,8)
	  
	  , @JobQty					BIGINT
	  , @matl_unit_cost_usd		DECIMAL(18,8)
	  , @matl_landed_cost_usd	DECIMAL(18,8)
	  , @pi_fg_process_usd		DECIMAL(18,8)
	  , @pi_resin_usd			DECIMAL(18,8)
	  , @pi_vend_cost_usd		DECIMAL(18,8)
	  , @pi_hidden_profit_usd	DECIMAL(18,8)
	  , @sf_lbr_cost_usd		DECIMAL(18,8)
	  , @sf_ovhd_cost_usd		DECIMAL(18,8)
	  , @fg_lbr_cost_usd		DECIMAL(18,8)
	  , @fg_ovhd_cost_usd		DECIMAL(18,8)
	  , @matl_unit_cost_php		DECIMAL(18,8)
	  , @matl_landed_cost_php	DECIMAL(18,8)
	  , @pi_fg_process_php		DECIMAL(18,8)
	  , @pi_resin_php			DECIMAL(18,8)
	  , @pi_vend_cost_php		DECIMAL(18,8)
	  , @pi_hidden_profit_php	DECIMAL(18,8)
	  , @sf_lbr_cost_php		DECIMAL(18,8)
	  , @sf_ovhd_cost_php		DECIMAL(18,8)
	  , @fg_lbr_cost_php		DECIMAL(18,8)
	  , @fg_ovhd_cost_php		DECIMAL(18,8)

	  , @ItemPricingCost		DECIMAL(18,8)
	  , @CurrCode				NVARCHAR(10)
	  , @ExchRate				ExchRateType
	    
	DECLARE @ItemCost AS TABLE (  
		Item			ItemType  
	  , LotNumber		LotType  
	  , LotQty			QtyUnitType  
	  , LotMatlCost		CostPrcType  
	  , LotLandedCost   CostPrcType  
	  , ItemPricing		CostPrcType  
	  , SafetyStockQty  QtyUnitType  
	  , ProductCode		ProductCodeType  
	)  
	  
	DECLARE @ItemLotCost AS TABLE (  
		item				ItemType  
	  , matl_cost			AmountType  
	  , landed_cost			AmountType  
	  , safety_stock_cost  AmountType  
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
	  
	SET @RMProdCode = 'RM-' + @ProdCode  
	SET @SFProdCode = (CASE @ProdCode  
		   WHEN 'PACKNG' THEN 'SA-PACK'  
		   WHEN 'CS-AT' THEN 'SA-CSAT'  
		   ELSE 'SA-' + @ProdCode  
		   END) 
		   
	DECLARE itemCrsr CURSOR FAST_FORWARD FOR  
	SELECT i.item  
	  , ll.lot  
	  , ll.qty_on_hand
	  , l.create_date
	FROM item AS i
		LEFT OUTER JOIN lot_loc AS ll
			ON i.item = ll.item  
		LEFT OUTER JOIN lot AS l
			ON ll.lot = l.lot
			  AND ll.item = l.item
	WHERE i.product_code = @RMProdCode 
		OR i.product_code = @SFProdCode  
	  
	OPEN itemCrsr  
	FETCH FROM itemCrsr INTO  
		@Item
	  , @LotNumber
	  , @LotQty
	  , @LotCreateDate
	  
	WHILE (@@FETCH_STATUS = 0)  
	BEGIN  
	   
		SELECT @SafetyStockQty = SUM(ISNULL(iw.qty_reorder, 0))  
		FROM itemwhse AS iw
			INNER JOIN whse AS w
				ON iw.whse = w.whse  
		WHERE iw.item = @Item AND w.dedicated_inventory = 0
		 
		SELECT TOP(1) @StdItemPrice = (unit_price1 * 1.2)
					, @CurrCode = curr_code
					  
		FROM itemprice
		WHERE item = @Item
		  AND effect_date <= @LotCreateDate
		ORDER BY effect_date DESC

		IF @CurrCode <> 'PHP'
		BEGIN
			EXEC dbo.LSP_CurrencyConversionModSp @LotCreateDate, @CurrCode, 'PHP', @StdItemPrice, @StdItemPrice_PHP OUTPUT, @ExchRate OUTPUT
		END
		ELSE
		BEGIN
			SET @StdItemPrice_PHP = @StdItemPrice
		END
		 
		  
		IF @Item LIKE 'SF-%'
		BEGIN
			
			IF EXISTS(SELECT * FROM job WHERE job = @LotNumber AND item = @Item)
			BEGIN			
				
				INSERT INTO #DMActualCost
				EXEC dbo.LSP_DM_ActlCost_GetJobMatlTransCostingSp @LotNumber, 0, @Item, @LotCreateDate, @LotQty
								
				INSERT INTO @ItemCost  
				SELECT @Item  
					 , @LotNumber  
					 , ISNULL(@LotQty,0)  
					 , (matl_unit_cost_php + pi_fg_process_php + pi_resin_php + pi_hidden_profit_php) / job_qty
					 , (matl_landed_cost_php / job_qty  )
					 , @StdItemPrice_PHP  
					 , @SafetyStockQty  
					 , @ProdCode
				FROM #DMActualCost
				WHERE [Level] = 0
								
			END
			ELSE
			BEGIN
			
				SELECT TOP(1) 
					   @ItemPricingCost = (unit_price1 / 0.9)
					 , @CurrCode = curr_code
					  
				FROM itemprice
				WHERE item = @Item
				  AND effect_date <= @LotCreateDate
				ORDER BY effect_date DESC
				
				EXEC dbo.LSP_CurrencyConversionModSp @LotCreateDate, @CurrCode, 'PHP', @ItemPricingCost, @matl_unit_cost_php OUTPUT, @ExchRate OUTPUT
				EXEC dbo.LSP_CurrencyConversionModSp @LotCreateDate, @CurrCode, 'USD', @ItemPricingCost, @matl_unit_cost_usd OUTPUT, @ExchRate OUTPUT
								
				INSERT INTO @ItemCost  
				SELECT @Item  
					 , @LotNumber  
					 , ISNULL(@LotQty,0)  
					 , @matl_unit_cost_php
					 , 0
					 , @StdItemPrice_PHP  
					 , @SafetyStockQty  
					 , @ProdCode
				
			END
		
		END
		ELSE
		BEGIN
		
			EXEC dbo.LSP_ActlCost_GetMatlCostingSp @Item, @LotNumber, @LotCreateDate
					  , @JobQty OUTPUT
					  , @matl_unit_cost_usd OUTPUT, @matl_landed_cost_usd OUTPUT
					  , @pi_fg_process_usd OUTPUT, @pi_resin_usd OUTPUT, @pi_vend_cost_usd OUTPUT, @pi_hidden_profit_usd OUTPUT
					  , @sf_lbr_cost_usd OUTPUT, @sf_ovhd_cost_usd OUTPUT
					  , @fg_lbr_cost_usd OUTPUT, @fg_ovhd_cost_usd OUTPUT
					  , @matl_unit_cost_php OUTPUT, @matl_landed_cost_php OUTPUT
					  , @pi_fg_process_php OUTPUT, @pi_resin_php OUTPUT, @pi_vend_cost_php OUTPUT, @pi_hidden_profit_php OUTPUT
					  , @sf_lbr_cost_php OUTPUT, @sf_ovhd_cost_php OUTPUT
					  , @fg_lbr_cost_php OUTPUT, @fg_ovhd_cost_php OUTPUT
			
			INSERT INTO @ItemCost  
			SELECT @Item  
				 , @LotNumber  
				 , ISNULL(@LotQty,0)  
				 , (@matl_unit_cost_php + @pi_fg_process_php + @pi_resin_php + @pi_hidden_profit_php)
				 , (@matl_landed_cost_php)
				 , @StdItemPrice_PHP  
				 , @SafetyStockQty  
				 , @ProdCode			
		END
				
		SELECT @JobQty = 0
			 , @matl_unit_cost_usd = 0
			 , @matl_landed_cost_usd = 0
			 , @pi_fg_process_usd = 0
			 , @pi_resin_usd = 0
			 , @pi_vend_cost_usd = 0
			 , @pi_hidden_profit_usd = 0
			 , @sf_lbr_cost_usd = 0
			 , @sf_ovhd_cost_usd = 0
			 , @fg_lbr_cost_usd = 0
			 , @fg_ovhd_cost_usd = 0
			 , @matl_unit_cost_php = 0
			 , @matl_landed_cost_php = 0
			 , @pi_fg_process_php = 0
			 , @pi_resin_php = 0
			 , @pi_vend_cost_php = 0
			 , @pi_hidden_profit_php = 0
			 , @sf_lbr_cost_php = 0
			 , @sf_ovhd_cost_php = 0
			 , @fg_lbr_cost_php = 0
			 , @fg_ovhd_cost_php = 0
			  
			 , @ItemPricingCost = 0
			 , @CurrCode = ''
			 , @ExchRate = 0
	  
	 FETCH NEXT FROM itemCrsr INTO  
		@Item
	  , @LotNumber
	  , @LotQty
	  , @LotCreateDate
	  
	END  
	  
	CLOSE itemCrsr  
	DEALLOCATE itemCrsr  
	  
	  
	INSERT INTO @ItemLotCost  
	SELECT Item  
	  , SUM(LotQty * LotMatlCost) AS TotalLotCost  
	  , SUM(LotQty * LotLandedCost) AS TotalLandedCost  
	  , (MAX(SafetyStockQty) * MAX(ItemPricing)) AS SafetyStockCost  
	  
	FROM @ItemCost  
	GROUP BY Item  
	  
	SELECT @InvtyMaterialCost = SUM(ISNULL(matl_cost, 0))
	  , @InvtyLandedCost = SUM(ISNULL(landed_cost, 0))
	  , @SafetyMaterialCost = SUM(ISNULL(safety_stock_cost, 0))
	--  , @SafetyLandedCost = SUM(safety_landed_cost)  
	  
	FROM @ItemLotCost  
	  
	
	--SELECT * FROM @ItemCost
	--SELECT * FROM @ItemLotCost
	  
	--SELECT @InvtyMaterialCost AS InvtyMatl
	--  , @InvtyLandedCost AS InvtyLanded
	--  , @SafetyMaterialCost AS SafetyMatl
  
--SELECT * FROM @ItemLotCost  
  
END