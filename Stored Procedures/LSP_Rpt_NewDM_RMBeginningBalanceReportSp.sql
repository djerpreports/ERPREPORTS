ALTER PROCEDURE LSP_Rpt_NewDM_RMBeginningBalanceReportSp (
--DECLARE
	@TransDate		DateType		--= '05/01/2021'
  , @ProdCode		ProductCodeType --= 'IPC'
) AS

BEGIN

	IF OBJECT_ID('tempdb..#DMActualCost') IS NOT NULL
		DROP TABLE #DMActualCost

	SET @TransDate = ISNULL(@TransDate, GETDATE())
	SET @ProdCode = ISNULL(NULLIF(@ProdCode, 'ALL'), '%')
	
	DECLARE  
		@Item					ItemType  
	  , @Description			DescriptionType  
	  , @VendNum				VendNumType  
	  , @VendName				NameType  
	  , @ProductCode			ProductCodeType  
	  , @QtyOnHand				QtyUnitType  
	  , @Location				LocType  
	  , @LotNo					LotType  
	  , @LotCreateDate			DateType  
	  , @U_m					UmType  
	  
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
	
	DECLARE @report_set AS TABLE (
		item				ItemType
	  , description			DescriptionType
	  , vend_num			VendNumType
	  , name				NameType
	  , product_code		ProductCodeType
	  , qty_on_hand			QtyUnitType
	  , loc					LocType
	  , lot					LotType
	  , lot_create_date		DateType
	  , u_m					UmType
	  , matl_unit_cost_php  CostPrcType
	  , landed_cost_php		CostPrcType
	  , resin_cost_php		CostPrcType
	  , pi_process_cost_php  CostPrcType
	  , pi_hidden_profit_php CostPrcType
	  , sf_added_value_php  CostPrcType
	  , matl_unit_cost_usd  CostPrcType
	  , landed_cost_usd   CostPrcType
	  , resin_cost_usd   CostPrcType
	  , pi_process_cost_usd  CostPrcType
	  , pi_hidden_profit_usd CostPrcType
	  , sf_added_value_usd  CostPrcType
	  , rm_cost_php    AmountType
	  , rm_cost_usd    AmountType	  
	)  

	IF @ProdCode = 'PACKNG'
	BEGIN
	 SET @ProdCode = 'PACK%'
	END
	
	IF @ProdCode <> 'PS-RM'
	BEGIN
		DECLARE materialCrsr CURSOR FAST_FORWARD FOR
		SELECT m.item
			 , i.description
			 , ISNULL(iv.vend_num, '')
			 , ISNULL(va.name, '')
			 , i.u_m
			 , i.product_code
			 , SUM(qty)
			 , m.loc
			 , m.lot
			 , l.create_date
		
		FROM matltran AS m
			JOIN item AS i
				ON m.item = i.item
			LEFT OUTER JOIN lot AS l
				ON m.item = l.item
					AND m.lot = l.lot
			LEFT OUTER JOIN itemvend AS iv
				ON iv.item = m.item
					AND iv.rank = 1
			LEFT OUTER JOIN vendaddr AS va
				ON iv.vend_num = va.vend_num
		 WHERE m.trans_date <= dbo.DayEndOf(@TransDate)
		  -- AND (i.product_code = 'PS-RM' OR i.product_code LIKE 'RM-%'  
		  --OR i.product_code LIKE 'SA-%' OR i.product_code = 'PI-RM')  
			AND ((i.product_code LIKE 'RM-' + @ProdCode) OR (i.product_code LIKE 'SA-' + @ProdCode))
		    AND (m.trans_type <> 'C' AND m.trans_type <> 'N')
		 GROUP BY m.item, i.description, iv.vend_num, va.name, i.u_m, i.product_code, m.loc, m.lot, l.create_date
		 HAVING SUM(qty) <> 0
		 ORDER BY m.item
	END
	ELSE
	BEGIN
		DECLARE materialCrsr CURSOR FAST_FORWARD FOR
		SELECT m.item
		     , i.description
		     , ISNULL(iv.vend_num, '')
		     , ISNULL(va.name, '')
		     , i.u_m
		     , i.product_code
		     , SUM(qty)
		     , m.loc
		     , m.lot
		     , l.create_date
		 
		 FROM matltran AS m
			JOIN item AS i
				ON m.item = i.item
			LEFT OUTER JOIN lot AS l
				ON m.item = l.item
					AND m.lot = l.lot
			LEFT OUTER JOIN itemvend AS iv
				ON iv.item = m.item
					AND iv.rank = 1
			LEFT OUTER JOIN vendaddr AS va
				ON iv.vend_num = va.vend_num
		 WHERE m.trans_date <= dbo.DayEndOf(@TransDate)
		  -- AND (i.product_code = 'PS-RM' OR i.product_code LIKE 'RM-%'
		  --OR i.product_code LIKE 'SA-%' OR i.product_code = 'PI-RM')
		  AND (i.product_code = @ProdCode)
		   AND (m.trans_type <> 'C' AND m.trans_type <> 'N')
		 GROUP BY m.item, i.description, iv.vend_num, va.name, i.u_m, i.product_code, m.loc, m.lot, l.create_date
		 HAVING SUM(qty) <> 0
		 ORDER BY m.item
	
	END
	
	OPEN materialCrsr
	FETCH FROM materialCrsr INTO
		@Item
	  , @Description
	  , @VendNum
	  , @VendName
	  , @U_m
	  , @ProductCode
	  , @QtyOnHand
	  , @Location
	  , @LotNo
	  , @LotCreateDate
	
	WHILE(@@FETCH_STATUS = 0)
	BEGIN
	IF @Item LIKE 'SF-%'
		BEGIN
			
			IF EXISTS(SELECT * FROM job WHERE job = @LotNo AND item = @Item)
			BEGIN
				SELECT @JobQty = qty_complete 
				FROM job
				WHERE job = @LotNo AND item = @Item
				
				TRUNCATE TABLE #DMActualCost
				
				INSERT INTO #DMActualCost
				EXEC dbo.LSP_DM_ActlCost_GetJobMatlTransCostingSp @LotNo, 0, @Item, @TransDate, @JobQty
						
				INSERT INTO @report_set
				SELECT @Item
				   , @Description
				   , @VendNum
				   , @VendName
				   , @ProductCode
				   , ISNULL(@QtyOnHand, 0)
				   , @Location
				   , @LotNo
				   , @LotCreateDate
				   , @U_m
				   , ISNULL(matl_unit_cost_php / job_qty, 0)
				   , ISNULL(matl_landed_cost_php / job_qty, 0)
				   , ISNULL(pi_resin_php / job_qty, 0)
				   , ISNULL(pi_fg_process_php / job_qty, 0)
				   , ISNULL(pi_hidden_profit_php / job_qty, 0)
				   , ISNULL((sf_lbr_cost_php + sf_ovhd_cost_php) / job_qty, 0)
				   
				   , ISNULL(matl_unit_cost_usd / job_qty, 0)
				   , ISNULL(matl_landed_cost_usd / job_qty, 0)
				   , ISNULL(pi_resin_usd / job_qty, 0)
				   , ISNULL(pi_fg_process_usd / job_qty, 0)
				   , ISNULL(pi_hidden_profit_usd / job_qty, 0)
				   , ISNULL((sf_lbr_cost_usd + sf_ovhd_cost_usd) / job_qty, 0)
				   
				   , ( ISNULL(matl_unit_cost_php, 0) + ISNULL(matl_landed_cost_php, 0)+ ISNULL(pi_resin_php, 0)
					    + ISNULL(pi_fg_process_php, 0)+ ISNULL(pi_hidden_profit_php, 0)+ ISNULL(sf_lbr_cost_php, 0)
					    + ISNULL(sf_ovhd_cost_php , 0) ) / job_qty
					 * ISNULL(@QtyOnHand, 0)

				   , ( ISNULL(matl_unit_cost_usd, 0) + ISNULL(matl_landed_cost_usd, 0)+ ISNULL(pi_resin_usd, 0)
					    + ISNULL(pi_fg_process_usd, 0)+ ISNULL(pi_hidden_profit_usd, 0)+ ISNULL(sf_lbr_cost_usd, 0)
					    + ISNULL(sf_ovhd_cost_usd , 0) / job_qty)
					 * ISNULL(@QtyOnHand, 0)
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
				  AND effect_date <= @TransDate
				ORDER BY effect_date DESC
				
				EXEC dbo.LSP_CurrencyConversionModSp @TransDate, @CurrCode, 'PHP', @ItemPricingCost, @matl_unit_cost_php OUTPUT, @ExchRate OUTPUT
				EXEC dbo.LSP_CurrencyConversionModSp @TransDate, @CurrCode, 'USD', @ItemPricingCost, @matl_unit_cost_usd OUTPUT, @ExchRate OUTPUT
								
				
				INSERT INTO @report_set
				SELECT @Item
				   , @Description
				   , @VendNum
				   , @VendName
				   , @ProductCode
				   , ISNULL(@QtyOnHand, 0)
				   , @Location
				   , @LotNo
				   , @LotCreateDate
				   , @U_m
				   , ISNULL(@matl_unit_cost_php, 0)
				   , 0
				   , 0
				   , 0
				   , 0
				   , 0
				   , ISNULL(@matl_unit_cost_usd, 0)
				   , 0
				   , 0
				   , 0
				   , 0
				   , 0				
				   , ISNULL(@matl_unit_cost_php, 0) * ISNULL(@QtyOnHand, 0)
				   , ISNULL(@matl_unit_cost_usd, 0) * ISNULL(@QtyOnHand, 0)
			END
		
		END
		ELSE
		BEGIN
					
			EXEC dbo.LSP_ActlCost_GetMatlCostingSp @Item, @LotNo, @TransDate
					  , @JobQty OUTPUT
					  , @matl_unit_cost_usd OUTPUT, @matl_landed_cost_usd OUTPUT
					  , @pi_fg_process_usd OUTPUT, @pi_resin_usd OUTPUT, @pi_vend_cost_usd OUTPUT, @pi_hidden_profit_usd OUTPUT
					  , @sf_lbr_cost_usd OUTPUT, @sf_ovhd_cost_usd OUTPUT
					  , @fg_lbr_cost_usd OUTPUT, @fg_ovhd_cost_usd OUTPUT
					  , @matl_unit_cost_php OUTPUT, @matl_landed_cost_php OUTPUT
					  , @pi_fg_process_php OUTPUT, @pi_resin_php OUTPUT, @pi_vend_cost_php OUTPUT, @pi_hidden_profit_php OUTPUT
					  , @sf_lbr_cost_php OUTPUT, @sf_ovhd_cost_php OUTPUT
					  , @fg_lbr_cost_php OUTPUT, @fg_ovhd_cost_php OUTPUT			
			
			INSERT INTO @report_set
			SELECT @Item
			   , @Description
			   , @VendNum
			   , @VendName
			   , @ProductCode
			   , ISNULL(@QtyOnHand, 0)
			   , @Location
			   , @LotNo
			   , @LotCreateDate
			   , @U_m
			   , ISNULL(@matl_unit_cost_php, 0)
			   , ISNULL(@matl_landed_cost_php, 0)
			   , ISNULL(@pi_resin_php, 0)
			   , ISNULL(@pi_fg_process_php, 0)
			   , ISNULL(@pi_hidden_profit_php, 0)
			   , ISNULL(@sf_lbr_cost_php, 0) + ISNULL(@sf_ovhd_cost_php, 0)
			   , ISNULL(@matl_unit_cost_usd, 0)
			   , ISNULL(@matl_landed_cost_usd, 0)
			   , ISNULL(@pi_resin_usd, 0)
			   , ISNULL(@pi_fg_process_usd, 0)
			   , ISNULL(@pi_hidden_profit_usd, 0)
			   , ISNULL(@sf_lbr_cost_usd, 0) + ISNULL(@sf_ovhd_cost_usd, 0)
			   , (ISNULL(@matl_unit_cost_php, 0) + ISNULL(@matl_landed_cost_php, 0) + ISNULL(@pi_resin_php, 0)
					+ ISNULL(@pi_fg_process_php, 0) + ISNULL(@pi_hidden_profit_php, 0)+ ISNULL(@sf_lbr_cost_php, 0) 
					+ ISNULL(@sf_ovhd_cost_php, 0) ) * ISNULL(@QtyOnHand, 0)
			   , (ISNULL(@matl_unit_cost_usd, 0) + ISNULL(@matl_landed_cost_usd, 0) + ISNULL(@pi_resin_usd, 0)
				    + ISNULL(@pi_fg_process_usd, 0) + ISNULL(@pi_hidden_profit_usd, 0) + ISNULL(@sf_lbr_cost_usd, 0) 
				    + ISNULL(@sf_ovhd_cost_usd, 0) ) * ISNULL(@QtyOnHand, 0)
			
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
	
		 FETCH NEXT FROM materialCrsr INTO
			 @Item
		   , @Description
		   , @VendNum
		   , @VendName
		   , @U_m
		   , @ProductCode
		   , @QtyOnHand
		   , @Location
		   , @LotNo
		   , @LotCreateDate
	
	END
	
	
	CLOSE materialCrsr
	DEALLOCATE materialCrsr

  
	SELECT *
	FROM @report_set

END