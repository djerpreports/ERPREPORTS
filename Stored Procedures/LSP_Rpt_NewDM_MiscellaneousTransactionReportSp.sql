--EXEC dbo.LSP_Rpt_NewDM_MiscellaneousTransactionReportSp '05/01/2020', '05/31/2020'

--ALTER PROCEDURE LSP_Rpt_NewDM_MiscellaneousTransactionReportSp (
DECLARE
	@StartDate					DateType	= '05/01/2020'
  , @EndDate					DateType	= '05/31/2020'
--) AS
BEGIN

	IF OBJECT_ID('tempdb..#DMActualCost') IS NOT NULL
		DROP TABLE #DMActualCost
	IF OBJECT_ID('tempdb..#MiscTransReport') IS NOT NULL
		DROP TABLE #MiscTransReport
	
	DECLARE
		@TransDate				DATETIME
	  , @TransType				NVARCHAR(2)
	  , @TransDesc				NVARCHAR(50)
	  , @JobOrLot				NVARCHAR(50)
	  , @Suffix					NVARCHAR(5)
	  , @Item					NVARCHAR(40)
	  , @ItemDesc				NVARCHAR(100)
	  , @QtyCompleted			DECIMAL(18,8)
	  , @QtyScrapped			DECIMAL(18,8)
	  , @OperNum				NVARCHAR(10)
	  , @Wc						NVARCHAR(50)
	  , @Employee				NVARCHAR(20)
	  , @MiscTransClass			NVARCHAR(5)
	  , @ReasonCode				NVARCHAR(20)
	  , @ReasonDesc				NVARCHAR(100)
	  , @TransQty				DECIMAL(18,8)
	  , @ABSTransQty			DECIMAL(18,8)
	  
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

	CREATE TABLE #MiscTransReport (
		TransDate				DATETIME
	  , TransType				NVARCHAR(2)
	  , TransDesc				NVARCHAR(50)
	  , SummaryGroup			NVARCHAR(50)
	  , JobOrLot				NVARCHAR(50)
	  , Suffix					NVARCHAR(5)
	  , Item					NVARCHAR(40)
	  , ItemDesc				NVARCHAR(100)
	  , QtyCompleted			DECIMAL(18,8)
	  , QtyScrapped				DECIMAL(18,8)
	  , OperNum					NVARCHAR(10)
	  , Wc						NVARCHAR(50)
	  , Employee				NVARCHAR(20)
	  , MiscTransClass			NVARCHAR(5)
	  , ReasonCode				NVARCHAR(20)
	  , ReasonDesc				NVARCHAR(100)
	  , TransQty				DECIMAL(18,8)
	  , MatlCost_PHP			DECIMAL(18,8)
	  , MatlLandedCost_PHP		DECIMAL(18,8)
	  , PIFGProcess_PHP			DECIMAL(18,8)
	  , PIResin_PHP				DECIMAL(18,8)
	  , PIHiddenProfit_PHP		DECIMAL(18,8)
	  , SFAddedCost_PHP			DECIMAL(18,8)	  
	  , FGAddedCost_PHP			DECIMAL(18,8)
	)

	SELECT @StartDate = dbo.MidnightOf(@StartDate)
		 , @EndDate = dbo.DayEndOf(@EndDate)

	DECLARE MiscTransCrsr CURSOR FAST_FORWARD FOR
	SELECT jt.trans_date AS TransDate
		 , 'G' As TransType
		 , 'SF Scrap Data' AS TransDesc
		 , j.job
		 , j.suffix
		 , j.item
		 , i.description
		 , jt.qty_complete AS QtyCompleted
		 , jt.qty_scrapped AS QtyScrapped
		 , jt.oper_num AS OperNum
		 , jt.wc
		 , jt.emp_num AS Employee
		 , 'S' AS MiscTransClass
		 , '' AS ReasonCode
		 , 'SF Scrap' AS ReasonDesc
		 , jt.qty_scrapped * (-1) AS TransQty
	
	FROM job AS j
		JOIN jobtran AS jt
			ON j.job = jt.job
			  AND j.suffix = j.suffix
		JOIN item AS i
			ON j.item = i.item
	WHERE jt.trans_date BETWEEN @StartDate and @EndDate
	  AND j.item LIKE 'SF-%'
	  AND jt.qty_scrapped > 0
	UNION ALL
	SELECT mt.trans_date AS TransDate
		 , mt.trans_type AS TransType
		 , CASE mt.trans_type
				WHEN 'G' THEN 'Miscellaneous Issue'
				WHEN 'H' THEN 'Miscellaneous Receipt'
				WHEN 'B' THEN 'Cycle Count'
				ELSE ''
		   END AS TransDesc
		 , mt.lot
		 , 0
		 , mt.item
		 , i.description
		 , NULL AS QtyCompleted
		 , NULL AS QtyScrapped
		 , NULL AS OperNum
		 , NULL AS wc
		 , NULL AS Employee
		 , CASE WHEN mt.trans_type = 'G'
					THEN CASE WHEN r1.description LIKE '%Scrap%' 
								THEN 'S'  
							  WHEN r1.description LIKE '%Request%' 
								THEN 'R'  
							  ELSE '' END       
				ELSE 'N'  
			END AS MiscTransClass
		 , mt.reason_code AS ReasonCode
		 , CASE mt.trans_type
				WHEN 'G' THEN r1.description
				WHEN 'H' THEN r2.description
				WHEN 'B' THEN 'Cycle Count'
				ELSE ''
		   END AS ReasonDesc
		 , mt.qty
	FROM matltran AS mt
		JOIN item AS i
			ON mt.item = i.item
		LEFT OUTER JOIN reason AS r1
			ON mt.reason_code = r1.reason_code
				AND r1.reason_class = 'MISC ISSUE'
		LEFT OUTER JOIN reason AS r2
			ON mt.reason_code = r2.reason_code
				AND r2.reason_class = 'MISC RCPT'
	WHERE mt.trans_date BETWEEN @StartDate AND @EndDate
	  AND mt.trans_type IN ('B','G','H')
	 -- AND mt.item = 'SF-PH017' AND mt.lot = '20HS-00015'
	/*****
	SELECT jt.trans_date AS TransDate
		 , 'G' As TransType
		 , 'Scrap Data' AS TransDesc
		 , j.job
		 , j.suffix
		 , j.item
		 , i.description
		 , jt.qty_complete AS QtyCompleted
		 , jt.qty_scrapped AS QtyScrapped
		 , jt.oper_num AS OperNum
		 , jt.wc
		 , jt.emp_num AS Employee
		 , 'S' AS MiscTransClass
		 , '' AS ReasonCode
		 , 'SF Scrap' AS ReasonDesc
		 , jt.qty_scrapped * (-1) AS TransQty
	
	FROM job AS j
		JOIN jobtran AS jt
			ON j.job = jt.job
			  AND j.suffix = j.suffix
		JOIN item AS i
			ON j.item = i.item
	WHERE jt.trans_date BETWEEN @StartDate and @EndDate
	  AND j.item LIKE 'SF-%'
	  AND jt.qty_scrapped > 0
	UNION ALL
	SELECT mt.trans_date AS TransDate
		 , mt.trans_type AS TransType
		 , CASE mt.trans_type
				WHEN 'G' THEN 'Miscellaneous Issue'
				WHEN 'H' THEN 'Miscellaneous Receipt'
				WHEN 'B' THEN 'Cycle Count'
				ELSE ''
		   END AS TransDesc
		 , mt.lot
		 , 0
		 , mt.item
		 , i.description
		 , NULL AS QtyCompleted
		 , NULL AS QtyScrapped
		 , NULL AS OperNum
		 , NULL AS wc
		 , NULL AS Employee
		 , CASE WHEN mt.trans_type = 'G'
					THEN CASE WHEN r1.description LIKE '%Scrap%' 
								THEN 'S'  
							  WHEN r1.description LIKE '%Request%' 
								THEN 'R'  
							  ELSE '' END       
				ELSE 'N'  
			END AS MiscTransClass
		 , mt.reason_code AS ReasonCode
		 , CASE mt.trans_type
				WHEN 'G' THEN r1.description
				WHEN 'H' THEN r2.description
				WHEN 'B' THEN 'Cycle Count'
				ELSE ''
		   END AS ReasonDesc
		 , mt.qty
	FROM matltran AS mt
		JOIN item AS i
			ON mt.item = i.item
		LEFT OUTER JOIN reason AS r1
			ON mt.reason_code = r1.reason_code
				AND r1.reason_class = 'MISC ISSUE'
		LEFT OUTER JOIN reason AS r2
			ON mt.reason_code = r2.reason_code
				AND r2.reason_class = 'MISC RCPT'
	WHERE mt.trans_date BETWEEN @StartDate AND @EndDate
	  AND mt.trans_type IN ('B','G','H')
	  *******/  
	
	OPEN MiscTransCrsr
	FETCH FROM MiscTransCrsr INTO
		@TransDate
	  , @TransType
	  , @TransDesc
	  , @JobOrLot
	  , @Suffix
	  , @Item
	  , @ItemDesc
	  , @QtyCompleted
	  , @QtyScrapped
	  , @OperNum
	  , @Wc
	  , @Employee
	  , @MiscTransClass
	  , @ReasonCode
	  , @ReasonDesc
	  , @TransQty
	
	WHILE @@FETCH_STATUS = 0
	BEGIN
	
		IF @Item LIKE 'SF-%'
		BEGIN
			
			IF EXISTS(SELECT * FROM job WHERE job = @JobOrLot AND item = @Item)
			BEGIN
				
				SET @ABSTransQty = ABS(@TransQty)				
				
				TRUNCATE TABLE #DMActualCost
				
				INSERT INTO #DMActualCost
				EXEC dbo.LSP_DM_ActlCost_GetJobMatlTransCostingSp @JobOrLot, 0, @Item, @TransDate, @ABSTransQty
						
				INSERT INTO #MiscTransReport
				SELECT @TransDate
					 , @TransType
					 , @TransDesc
					 , CASE WHEN @MiscTransClass = 'S'
									THEN 'Breakdown of Scrap'
							WHEN @MiscTransClass = 'R'
									THEN 'Breakdown of Request'
							ELSE @TransDesc END
					 , @JobOrLot
					 , @Suffix
					 , @Item
					 , @ItemDesc
					 , @QtyCompleted
					 , @QtyScrapped
					 , @OperNum
					 , @Wc
					 , @Employee
					 , @MiscTransClass
					 , @ReasonCode
					 , @ReasonDesc
					 , @TransQty
					 , matl_unit_cost_php / job_qty
					 , matl_landed_cost_php / job_qty
					 , pi_fg_process_php / job_qty
					 , pi_resin_php / job_qty
					 , pi_hidden_profit_php / job_qty
					 , (sf_lbr_cost_php + sf_ovhd_cost_php) / job_qty
					 , (fg_lbr_cost_php + fg_ovhd_cost_php) / job_qty
					 
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
								
				
				INSERT INTO #MiscTransReport
				SELECT @TransDate
					 , @TransType
					 , @TransDesc
					 , CASE WHEN @MiscTransClass = 'S'
									THEN 'Breakdown of Scrap'
							WHEN @MiscTransClass = 'R'
									THEN 'Breakdown of Request'
							ELSE @TransDesc END
					 , @JobOrLot
					 , @Suffix
					 , @Item
					 , @ItemDesc
					 , @QtyCompleted
					 , @QtyScrapped
					 , @OperNum
					 , @Wc
					 , @Employee
					 , @MiscTransClass
					 , @ReasonCode
					 , @ReasonDesc
					 , @TransQty
					 , ISNULL(@matl_unit_cost_php, 0)
					 , 0
					 , 0
					 , 0
					 , 0
					 , 0
					 , 0
					 
				
			END
		
		END
		ELSE
		BEGIN
					
			EXEC dbo.LSP_ActlCost_GetMatlCostingSp @Item, @JobOrLot, @TransDate
					  , @JobQty OUTPUT
					  , @matl_unit_cost_usd OUTPUT, @matl_landed_cost_usd OUTPUT
					  , @pi_fg_process_usd OUTPUT, @pi_resin_usd OUTPUT, @pi_vend_cost_usd OUTPUT, @pi_hidden_profit_usd OUTPUT
					  , @sf_lbr_cost_usd OUTPUT, @sf_ovhd_cost_usd OUTPUT
					  , @fg_lbr_cost_usd OUTPUT, @fg_ovhd_cost_usd OUTPUT
					  , @matl_unit_cost_php OUTPUT, @matl_landed_cost_php OUTPUT
					  , @pi_fg_process_php OUTPUT, @pi_resin_php OUTPUT, @pi_vend_cost_php OUTPUT, @pi_hidden_profit_php OUTPUT
					  , @sf_lbr_cost_php OUTPUT, @sf_ovhd_cost_php OUTPUT
					  , @fg_lbr_cost_php OUTPUT, @fg_ovhd_cost_php OUTPUT
			
			
			INSERT INTO #MiscTransReport
			SELECT @TransDate
				 , @TransType
				 , @TransDesc
				 , CASE WHEN @MiscTransClass = 'S'
									THEN 'Breakdown of Scrap'
							WHEN @MiscTransClass = 'R'
									THEN 'Breakdown of Request'
							ELSE @TransDesc END
				 , @JobOrLot
				 , @Suffix
				 , @Item
				 , @ItemDesc
				 , @QtyCompleted
				 , @QtyScrapped
				 , @OperNum
				 , @Wc
				 , @Employee
				 , @MiscTransClass
				 , @ReasonCode
				 , @ReasonDesc
				 , @TransQty
				 , ISNULL(@matl_unit_cost_php, 0)
				 , ISNULL(@matl_landed_cost_php, 0)
				 , ISNULL(@pi_fg_process_php, 0)
				 , ISNULL(@pi_resin_php, 0)
				 , ISNULL(@pi_hidden_profit_php, 0)
				 , ISNULL(@sf_lbr_cost_php, 0) + ISNULL(@sf_ovhd_cost_php, 0)
				 , ISNULL(@fg_lbr_cost_php, 0) + ISNULL(@fg_ovhd_cost_php, 0)
			
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
	
		FETCH NEXT FROM MiscTransCrsr INTO
			@TransDate
		  , @TransType
		  , @TransDesc
		  , @JobOrLot
		  , @Suffix
		  , @Item
		  , @ItemDesc
		  , @QtyCompleted
		  , @QtyScrapped
		  , @OperNum
		  , @Wc
		  , @Employee
		  , @MiscTransClass
		  , @ReasonCode
		  , @ReasonDesc
		  , @TransQty
	
	END
	
	CLOSE MiscTransCrsr
	DEALLOCATE MiscTransCrsr
		
	
	SELECT *
		 , TransQty * (MatlCost_PHP + MatlLandedCost_PHP 
						+ PIFGProcess_PHP + PIResin_PHP + PIHiddenProfit_PHP 
						+ SFAddedCost_PHP + FGAddedCost_PHP)
		   AS TotalCost_PHP
	
	FROM #MiscTransReport
	UNION ALL
	SELECT @EndDate
		 , 'G'
		 , 'Miscellaneous Issue'
		 , 'Miscellaneous Issue'
		 , ''
		 , ''
		 , 'Scrap Item'
		 , 'Scrap Description'
		 , NULL
		 , NULL
		 , NULL
		 , NULL
		 , NULL
		 , ''
		 , 'SCRAP'
		 , 'Scrap'
		 , 1 AS TransQty
		 , SUM(MatlCost_PHP * TransQty) 
		 , SUM(MatlLandedCost_PHP * TransQty) 
		 , SUM(PIFGProcess_PHP * TransQty) 
		 , SUM(PIResin_PHP * TransQty) 
		 , SUM(PIHiddenProfit_PHP * TransQty) 
		 , SUM(SFAddedCost_PHP * TransQty) 
		 , SUM(FGAddedCost_PHP * TransQty)
		 , SUM(TransQty * (MatlCost_PHP + MatlLandedCost_PHP 
						+ PIFGProcess_PHP + PIResin_PHP + PIHiddenProfit_PHP 
						+ SFAddedCost_PHP + FGAddedCost_PHP))
		   AS TotalCost_PHP	
	FROM #MiscTransReport
	WHERE TransDesc = 'Miscellaneous Issue'
	  AND ReasonDesc LIKE '%Scrap%'
	UNION ALL 
	SELECT @EndDate
		 , 'G'
		 , 'Miscellaneous Issue'
		 , 'Miscellaneous Issue'
		 , ''
		 , ''
		 , 'Request Item'
		 , 'Request Description'
		 , NULL
		 , NULL
		 , NULL
		 , NULL
		 , NULL
		 , ''
		 , 'REQ'
		 , 'Section Requests'
		 , 1 AS TransQty
		 , SUM(MatlCost_PHP * TransQty) 
		 , SUM(MatlLandedCost_PHP * TransQty) 
		 , SUM(PIFGProcess_PHP * TransQty) 
		 , SUM(PIResin_PHP * TransQty) 
		 , SUM(PIHiddenProfit_PHP * TransQty) 
		 , SUM(SFAddedCost_PHP * TransQty) 
		 , SUM(FGAddedCost_PHP * TransQty)
		 , SUM(TransQty * (MatlCost_PHP + MatlLandedCost_PHP 
						+ PIFGProcess_PHP + PIResin_PHP + PIHiddenProfit_PHP 
						+ SFAddedCost_PHP + FGAddedCost_PHP))
		   AS TotalCost_PHP	
	FROM #MiscTransReport
	WHERE TransDesc = 'Miscellaneous Issue'
	  AND ReasonDesc LIKE '%Request%'
END