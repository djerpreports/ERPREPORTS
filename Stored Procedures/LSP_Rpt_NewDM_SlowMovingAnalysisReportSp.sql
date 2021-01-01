--EXEC dbo.LSP_Rpt_NewDM_SlowMovingAnalysisReportSp 12

ALTER PROCEDURE LSP_Rpt_NewDM_SlowMovingAnalysisReportSp (
--DECLARE
	@Months					INT	--= 12
) AS
BEGIN

	IF OBJECT_ID('tempdb..#itemSM') IS NOT NULL
		DROP TABLE #itemSM
	IF OBJECT_ID('tempdb..#LatestPORcvdDate') IS NOT NULL
		DROP TABLE #LatestPORcvdDate 
	IF OBJECT_ID('tempdb..#LatestIssueDate') IS NOT NULL
		DROP TABLE #LatestIssueDate
	IF OBJECT_ID('tempdb..#ItemLotLocCosts') IS NOT NULL
		DROP TABLE #ItemLotLocCosts

	DECLARE   
		@StartDate				DateType
	  , @EndDate				DateType
	  , @Item					ItemType
	  , @Lot					LotType
	  , @QtyOnHand				QtyUnitType
	  , @LotCreateDate			DateType
	  
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

	  , @Remarks				NVARCHAR(50)
	  , @ItemPricingCost		DECIMAL(18,8)
	  , @CurrCode				NVARCHAR(10)
	  , @ExchRate				ExchRateType
	  
	CREATE TABLE #ItemLotLocCosts (
		item						NVARCHAR(60)
	  , lot_no						NVARCHAR(30)
	  , lot_createDate				DateTime
	  , qty_on_hand					DECIMAL(18,8)
	  , matl_cost_php				DECIMAL(18,8)
	  , matl_landed_cost_php		DECIMAL(18,8)
	  , pi_fg_process_php			DECIMAL(18,8)
	  , pi_resin_php				DECIMAL(18,8)
	  , pi_vend_cost_php			DECIMAL(18,8)
	  , pi_hidden_profit_php		DECIMAL(18,8)
	  , sf_lbr_cost_php				DECIMAL(18,8)
	  , sf_ovhd_cost_php			DECIMAL(18,8)
	  , ItemRemarks					NVARCHAR(50)
	)
	  
	SET @Months = ISNULL(NULLIF(@Months, 0),12)

	SELECT @StartDate = DATEADD(S, 0, DATEADD(M, DATEDIFF(m, 0, GETDATE())- @Months,0))  
		 , @EndDate =  DATEADD(S, -1, DATEADD(mm, DATEDIFF(m, 0, GETDATE()),0))  

	SELECT i.item  
		 , i.description  
		 , i.Uf_location  
		 , i.stat  
		 , i.product_code  
		, SUM(ISNULL(iw.qty_on_hand,0)) AS qty_on_hand
	INTO #itemSM
	FROM item AS i
		LEFT OUTER JOIN itemwhse AS iw
			ON i.item = iw.item		
	WHERE i.item NOT IN (SELECT m.item  	
						 FROM matltran AS m
							 JOIN item AS i2
								ON m.item = i2.item  
						 WHERE m.trans_type = 'I'  
						   AND m.trans_date BETWEEN @StartDate AND @EndDate  
						   AND i2.item NOT LIKE 'FG-%'  
						   AND (i2.product_code LIKE 'RM-%'  
								OR i2.product_code LIKE 'SF-%'  
								OR i2.product_code LIKE 'SC-%'  
								OR i2.product_code LIKE 'PI-RM-%'  
								OR i2.product_code LIKE 'URE-RM-%')  
						   AND i2.stat <> 'OBS')  
	  AND i.item NOT LIKE 'FG-%'  
	  AND i.stat <> 'O'  
	  AND (i.product_code NOT LIKE 'OS-%'  
	  AND i.product_code NOT LIKE '%-SUP')  
	GROUP BY i.item, i.description, i.Uf_location, i.stat, i.product_code

	SELECT m.item
		 , MAX(m.trans_date) AS issue_date
	INTO #LatestIssueDate
	FROM matltran AS m
		JOIN #itemSM AS i
			ON m.item = i.item
	WHERE m.trans_type = 'I'  
	  AND m.ref_type = 'J'
	GROUP BY m.item
	  
	SELECT m.item
		 , MAX(m.trans_date) AS latest_po_date
	INTO #LatestPORcvdDate
	FROM matltran AS m
		JOIN #itemSM AS i
			ON m.item = i.item
	WHERE m.trans_type = 'R'  
	  AND m.ref_type = 'P'
	GROUP BY m.item

	DECLARE itemLotCrsr CURSOR FAST_FORWARD FOR	
	SELECT sm.item
		 , l.lot
		 , l.qty_on_hand
		 , l.CreateDate
	FROM #itemSM AS sm
		JOIN lot_loc AS l
			ON sm.item = l.item
	--WHERE sm.item LIKE 'RM-%'
	--UNION
	--SELECT TOP(10) sm.item
	--	 , l.lot
	--	 , l.qty_on_hand
	--	 , l.CreateDate
	--FROM #itemSM AS sm
	--	JOIN lot_loc AS l
	--		ON sm.item = l.item
	--WHERE sm.item LIKE 'SF-%' AND l.lot NOT LIKE '%stock%'
	--UNION
	--SELECT TOP(10) sm.item
	--	 , l.lot
	--	 , l.qty_on_hand
	--	 , l.CreateDate
	--FROM #itemSM AS sm
	--	JOIN lot_loc AS l
	--		ON sm.item = l.item
	--WHERE sm.item LIKE 'SF-%' AND l.lot LIKE '%stock%'
			
	OPEN itemLotCrsr
	FETCH FROM itemLotCrsr INTO
		@Item
	  , @Lot
	  , @QtyOnHand
	  , @LotCreateDate
	  
	WHILE (@@FETCH_STATUS = 0)
	BEGIN

		IF @Item LIKE 'SF-%'
		BEGIN
			
			IF EXISTS(SELECT * FROM job WHERE job = @Lot)
			BEGIN			
				EXEC dbo.LSP_DM_ActlCost_GetJobMatlTransCostingSp @Lot, 0, @Item, @LotCreateDate, @QtyOnHand
				
				EXEC dbo.LSP_GetSlowMovingAnalysisReportRemarks @item, @Remarks OUTPUT
				
				INSERT INTO #ItemLotLocCosts
				SELECT @item
					 , @Lot
					 , @LotCreateDate
					 , @QtyOnHand
					 , matl_unit_cost_php / job_qty
					 , matl_landed_cost_php / job_qty
					 , pi_fg_process_php / job_qty
					 , pi_resin_php / job_qty
					 , pi_vend_cost_php / job_qty
					 , pi_hidden_profit_php / job_qty
					 , sf_lbr_cost_php / job_qty
					 , sf_ovhd_cost_php / job_qty
					 , @Remarks
				FROM ##ActualCost
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
								
				EXEC dbo.LSP_GetSlowMovingAnalysisReportRemarks @item, @Remarks OUTPUT
				
				INSERT INTO #ItemLotLocCosts
				SELECT @item
					 , @Lot
					 , @LotCreateDate
					 , @QtyOnHand
					 , @matl_unit_cost_php
					 , 0
					 , 0
					 , 0
					 , 0
					 , 0
					 , 0
					 , 0
					 , @Remarks
				
			END
		
		END
		ELSE
		BEGIN
		
			EXEC dbo.LSP_ActlCost_GetMatlCostingSp @Item, @Lot, @LotCreateDate
					  , @JobQty OUTPUT
					  , @matl_unit_cost_usd OUTPUT, @matl_landed_cost_usd OUTPUT
					  , @pi_fg_process_usd OUTPUT, @pi_resin_usd OUTPUT, @pi_vend_cost_usd OUTPUT, @pi_hidden_profit_usd OUTPUT
					  , @sf_lbr_cost_usd OUTPUT, @sf_ovhd_cost_usd OUTPUT
					  , @fg_lbr_cost_usd OUTPUT, @fg_ovhd_cost_usd OUTPUT
					  , @matl_unit_cost_php OUTPUT, @matl_landed_cost_php OUTPUT
					  , @pi_fg_process_php OUTPUT, @pi_resin_php OUTPUT, @pi_vend_cost_php OUTPUT, @pi_hidden_profit_php OUTPUT
					  , @sf_lbr_cost_php OUTPUT, @sf_ovhd_cost_php OUTPUT
					  , @fg_lbr_cost_php OUTPUT, @fg_ovhd_cost_php OUTPUT
			
			EXEC dbo.LSP_GetSlowMovingAnalysisReportRemarks @item, @Remarks OUTPUT
			
			INSERT INTO #ItemLotLocCosts
			SELECT @item
				 , @Lot
				 , @LotCreateDate
				 , @QtyOnHand
				 , ISNULL(@matl_unit_cost_php, 0)
				 , ISNULL(@matl_landed_cost_php, 0)
				 , ISNULL(@pi_fg_process_php, 0)
				 , ISNULL(@pi_resin_php, 0)
				 , ISNULL(@pi_vend_cost_php, 0)
				 , ISNULL(@pi_hidden_profit_php, 0)
				 , ISNULL(@sf_lbr_cost_php, 0)
				 , ISNULL(@sf_ovhd_cost_php, 0)
				 , ISNULL(@Remarks, '')
			
		END
		
		SET @Remarks = ''

		FETCH NEXT FROM itemLotCrsr INTO
			@Item
		  , @Lot
		  , @QtyOnHand
		  , @LotCreateDate
	
	END
	
	CLOSE itemLotCrsr
	DEALLOCATE itemLotCrsr

	--SELECT * FROM #itemSM WHERE item = 'SF-1G20081'
	
	SELECT sm.item
		 , sm.description
		 , sm.Uf_location
		 , CASE sm.stat
				WHEN 'A' THEN 'Active'  
				WHEN 'O' THEN 'Obsolete'  
				WHEN 'S' THEN 'Slow Moving' 
				ELSE sm.stat
		   END AS matl_stat
		 , sm.product_code
		 , ISNULL(SUM(ll.qty_on_hand),0) AS QtyOnHand
		 , ISNULL(SUM(ll.qty_on_hand * ll.matl_cost_php),0) AS TotalMatlCostPHP
		 , ISNULL(SUM(ll.qty_on_hand * ll.matl_landed_cost_php),0) AS TotalLandedCostPHP
		 , ISNULL(SUM(ll.qty_on_hand * ll.pi_fg_process_php),0) AS TotalPIFGProcessCostPHP
		 , ISNULL(SUM(ll.qty_on_hand * ll.pi_resin_php),0) AS TotalPIResinCostPHP
		 , ISNULL(SUM(ll.qty_on_hand * ll.pi_hidden_profit_php),0) AS TotalPIHiddenPHP
		 , ISNULL(SUM(ll.qty_on_hand * (ll.sf_lbr_cost_php + ll.sf_ovhd_cost_php )),0) AS TotalSFLbrCostPHP
		 , ISNULL(SUM(ll.qty_on_hand 
				* ( ll.matl_cost_php + ll.matl_landed_cost_php 
					 + ll.pi_fg_process_php + ll.pi_resin_php + ll.pi_hidden_profit_php
					 + ll.sf_lbr_cost_php + ll.sf_ovhd_cost_php )),0) AS TotalCostPHP
		 , r.latest_po_date AS LatestPODate
		 , i.issue_date AS LatestIssueDate
		 , ISNULL(ll.ItemRemarks, '') AS ItemRemarks
	FROM #itemSM AS sm
		LEFT OUTER JOIN #LatestPORcvdDate AS r
			ON sm.item = r.item
		LEFT OUTER JOIN #LatestIssueDate AS i
			ON sm.item = i.item
		LEFT OUTER JOIN #ItemLotLocCosts AS ll
			ON sm.item = ll.item
	GROUP BY sm.item
		 , sm.description
		 , sm.Uf_location
		 , sm.stat
		 , sm.product_code
		 , r.latest_po_date
		 , i.issue_date
		 , ll.ItemRemarks

END