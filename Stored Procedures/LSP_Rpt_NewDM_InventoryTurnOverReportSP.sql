ALTER PROCEDURE LSP_Rpt_NewDM_InventoryTurnOverReportSP (
--DECLARE
	@IsShowDetail					BIT				--=			1
  , @StartDate						DateType		OUTPUT
  , @EndDate						DateType		OUTPUT
) AS
BEGIN

	IF OBJECT_ID('tempdb..#InvtyTurnOverDtl') IS NOT NULL
		DROP TABLE #InvtyTurnOverDtl
	IF OBJECT_ID('tempdb..#DMActualCost') IS NOT NULL
		DROP TABLE #DMActualCost

	DECLARE  
		/*@StartDate					DateType
	  , @EndDate					DateType
	  , */@TransDate					DATETIME
	  , @TransType					NVARCHAR(25)  
	  , @ReasonCode					NVARCHAR(10)
	  , @ReasonDesc					NVARCHAR(80)
	  , @Qty						DECIMAL(18,8)
	  , @Item						NVARCHAR(60)
	  , @ItemDesc					NVARCHAR(80)
	  , @ProductCode				NVARCHAR(20)
	  , @LotNumber					NVARCHAR(30)
	  , @RefNum						NVARCHAR(40)
	  , @RefLine					INT
	  
	  , @JobQty						BIGINT
	  , @matl_unit_cost_usd			DECIMAL(18,8)
	  , @matl_landed_cost_usd		DECIMAL(18,8)
	  , @pi_fg_process_usd			DECIMAL(18,8)
	  , @pi_resin_usd				DECIMAL(18,8)
	  , @pi_vend_cost_usd			DECIMAL(18,8)
	  , @pi_hidden_profit_usd		DECIMAL(18,8)
	  , @sf_lbr_cost_usd			DECIMAL(18,8)
	  , @sf_ovhd_cost_usd			DECIMAL(18,8)
	  , @fg_lbr_cost_usd			DECIMAL(18,8)
	  , @fg_ovhd_cost_usd			DECIMAL(18,8)
	  , @matl_unit_cost_php			DECIMAL(18,8)
	  , @matl_landed_cost_php		DECIMAL(18,8)
	  , @pi_fg_process_php			DECIMAL(18,8)
	  , @pi_resin_php				DECIMAL(18,8)
	  , @pi_vend_cost_php			DECIMAL(18,8)
	  , @pi_hidden_profit_php		DECIMAL(18,8)
	  , @sf_lbr_cost_php			DECIMAL(18,8)
	  , @sf_ovhd_cost_php			DECIMAL(18,8)
	  , @fg_lbr_cost_php			DECIMAL(18,8)
	  , @fg_ovhd_cost_php			DECIMAL(18,8)

	  , @ItemPricingCost			DECIMAL(18,8)
	  , @CurrCode					NVARCHAR(10)
	  , @ExchRate					ExchRateType
	  
	  , @TransDateCrsr				DATETIME
	  , @ProdCodeCrsr				NVARCHAR(20)
	  , @ProdCode					NVARCHAR(20)
	  , @UsageQty					DECIMAL(18,8)
	  , @MatlUsage					DECIMAL(18,8)
	  , @LandedUsage				DECIMAL(18,8)
	  , @InvtyMatlCost				DECIMAL(18,8)
	  , @InvtyLandedCost			DECIMAL(18,8)
	  , @SafetyMatlCost				DECIMAL(18,8)
	
	  , @MonthGroup			NVARCHAR(200) = ''
	  , @MonthGroup2		NVARCHAR(200) = ''
	  , @Select1			NVARCHAR(1000) = ''
	  , @Select2			NVARCHAR(1000) = ''
	  , @CurrDate			DateType
	  , @SqlStr				NVARCHAR(4000)
	  , @ParmDefinition		NVARCHAR(500)
	  
	  ,	@DMInvty			DMInvtyTurnOverType
	  
	DECLARE @report_set AS TABLE (  
		trans_date					DateType
	  , trans_type					NVARCHAR(25)
	  , reason_code					ReasonCodeType
	  , reason_desc					DescriptionType
	  , qty							QtyUnitType
	  , usage_matl					AmountType
	  , usage_landed				AmountType	  
	  , usage_M1					AmountType
	  , usage_M2					AmountType
	  , usage_M3					AmountType
	  , usage_M4					AmountType
	  , usage_M5					AmountType
	  , usage_M6					AmountType
	  , usage_M7					AmountType
	  , usage_M8					AmountType
	  , usage_M9					AmountType
	  , usage_M10					AmountType
	  , usage_M11					AmountType
	  , usage_M12					AmountType
	  , usage_L1					AmountType
	  , usage_L2					AmountType
	  , usage_L3					AmountType
	  , usage_L4					AmountType
	  , usage_L5					AmountType
	  , usage_L6					AmountType
	  , usage_L7					AmountType
	  , usage_L8					AmountType
	  , usage_L9					AmountType
	  , usage_L10					AmountType
	  , usage_L11					AmountType
	  , usage_L12					AmountType
	  , item					    ItemType  
	  , item_desc					DescriptionType  
	  , product_code				ProductCodeType  
	  , lot							LotType  
	  , ref_num					    EmpJobCoPoRmaProjPsTrnNumType  
	  , ref_line					CoLineSuffixPoLineProjTaskRmaTrnLineType  
	  , invty_matl_cost				AmountType  
	  , invty_landed_cost			AmountType  
	  , safety_matl_cost			AmountType  
	  , report_group				NVARCHAR(10)  
	  , M1							AmountType  
	  , L1							AmountType  
	  , M2							AmountType  
	  , L2							AmountType  
	  , M3							AmountType  
	  , L3							AmountType  
	  , M4							AmountType  
	  , L4							AmountType  
	  , M5							AmountType  
	  , L5							AmountType  
	  , M6							AmountType  
	  , L6							AmountType  
	  , M7							AmountType  
	  , L7							AmountType  
	  , M8							AmountType  
	  , L8							AmountType  
	  , M9							AmountType  
	  , L9							AmountType  
	  , M10							AmountType  
	  , L10							AmountType  
	  , MAX_3Months					AmountType  
	  , L_MAX_3Months				AmountType  
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
	
	CREATE TABLE #InvtyTurnOverDtl(	
		TransDate					DATETIME
	  , TransType					NVARCHAR(25)  
	  , ReasonCode					NVARCHAR(10)
	  , ReasonDesc					NVARCHAR(80)
	  , Qty							DECIMAL(18,8)
	  , Item						NVARCHAR(60)
	  , ItemDesc					NVARCHAR(80)
	  , ProductCode					NVARCHAR(20)
	  , LotNumber					NVARCHAR(30)
	  , RefNum						NVARCHAR(40)
	  , RefLine						INT  
	  , MatlCost					DECIMAL(18,8)
	  , LandedCost					DECIMAL(18,8)
	)
	
	SELECT @StartDate =	DATEADD(S, 0, DATEADD(M,DATEDIFF(m, 0, GETDATE())-12,0))
		 , @EndDate = DATEADD(S, -1, DATEADD(mm, DATEDIFF(m, 0, GETDATE()),0))
	--/*****************
	--INSERT INTO #InvtyTurnOverDtl (TransDate, TransType, ReasonCode, ReasonDesc, Qty, Item, ItemDesc, ProductCode, LotNumber, RefNum, RefLine)
	DECLARE transCrsr CURSOR FAST_FORWARD FOR	
	SELECT mt.trans_date
		 , 'Ship'
		 , mt.reason_code
		 , NULL
		 , mt.qty * (-1)
		 , mt.item
		 , i.description
		 , CASE i.product_code
				WHEN 'SA-PACK'
					THEN 'SA-PACKNG'
				WHEN 'SA-CSAT'
					THEN 'SA-CS-AT'
				ELSE i.product_code
		   END
		 , mt.lot
		 , mt.ref_num
		 , mt.ref_line_suf
	FROM matltran AS mt
		LEFT OUTER JOIN item AS i
			ON mt.item = i.item
	WHERE mt.trans_date BETWEEN dbo.MidnightOf(@StartDate) AND dbo.DayEndOf(@EndDate)
	  AND (mt.trans_type = 'S'
			 AND (mt.ref_num NOT LIKE '%PDN%'
					AND mt.ref_num NOT LIKE '%PIS%'
					AND mt.ref_num NOT LIKE '%QAC%'
					AND mt.ref_num NOT LIKE '%PDE%'
					AND mt.ref_num NOT LIKE '%MDE%') )
	  AND (i.product_code NOT LIKE 'FG-%'
			AND i.product_code NOT LIKE 'OS-%'
			AND i.product_code NOT LIKE '%-SUP')
	UNION
	SELECT mt.trans_date
		 , CASE mt.trans_type
				WHEN 'I'
					THEN 'Issuance'
				WHEN 'W'
					THEN 'Withdrawal'
				ELSE NULL
		   END
		 , NULL
		 , NULL
		 , mt.qty * (-1)
		 , mt.item
		 , i.description
		 , CASE i.product_code
				WHEN 'SA-PACK'
					THEN 'SA-PACKNG'
				WHEN 'SA-CSAT'
					THEN 'SA-CS-AT'
				ELSE i.product_code  
			END
		 , mt.lot
		 , mt.ref_num
		 , mt.ref_line_suf
	FROM matltran AS mt
		LEFT OUTER JOIN item AS i
			ON mt.item = i.item
	WHERE mt.trans_date BETWEEN dbo.MidnightOf(@StartDate) AND dbo.DayEndOf(@EndDate)
	 AND (mt.trans_type = 'I'
			OR mt.trans_type = 'W')
	 AND (mt.ref_num LIKE '[1-9]_%'				--TO CHANGE [1-1] ==> [1-9]
			OR mt.ref_num LIKE '[1-9]_RM-%')	--TO CHANGE [1-1] ==> [1-9]
	 AND (i.product_code NOT LIKE 'FG-%'
			AND i.product_code NOT LIKE 'OS-%'
			AND i.product_code NOT LIKE '%-SUP' )
	UNION
	SELECT mt.trans_date  
		 , 'Misc. Issue'  
		 , mt.reason_code  
		 , r.description  
		 , mt.qty * (-1)  
		 , mt.item  
		 , i.description  
		 , CASE i.product_code  
				WHEN 'SA-PACK' 
					THEN 'SA-PACKNG'  
				WHEN 'SA-CSAT' 
					THEN 'SA-CS-AT'  
				ELSE i.product_code  
		   END  
		 , mt.lot  
		 , mt.ref_num  
		 , mt.ref_line_suf  
	FROM matltran AS mt
		LEFT OUTER JOIN item AS i
			ON mt.item = i.item 
		LEFT OUTER JOIN reason AS r
			ON mt.reason_code = r.reason_code 
			  AND r.reason_class = 'MISC ISSUE'  
	WHERE mt.trans_date BETWEEN dbo.MidnightOf(@StartDate) AND dbo.DayEndOf(@EndDate)  
	  AND (mt.trans_type = 'G'  
		  AND (mt.reason_code = 'ARM' 
				OR mt.reason_code = 'ETS'  
				OR mt.reason_code = 'PIU' 
				OR mt.reason_code = 'URE')  
			   )  
	  AND (i.product_code NOT LIKE 'FG-%'  
			AND i.product_code NOT LIKE 'OS-%'  
			AND i.product_code NOT LIKE '%-SUP')  

	OPEN transCrsr
	FETCH FROM transCrsr INTO
		@TransDate
	  , @TransType
	  , @ReasonCode
	  , @ReasonDesc
	  , @Qty
	  , @Item
	  , @ItemDesc
	  , @ProductCode
	  , @LotNumber
	  , @RefNum
	  , @RefLine	
	
	WHILE @@FETCH_STATUS = 0
	BEGIN
		
		IF @Item LIKE 'SF-%'
		BEGIN
			
			IF EXISTS(SELECT * FROM job WHERE job = @LotNumber AND item = @Item)
			BEGIN			
				TRUNCATE TABLE #DMActualCost
				
				INSERT INTO #DMActualCost
				EXEC dbo.LSP_DM_ActlCost_GetJobMatlTransCostingSp @LotNumber, 0, @Item, @TransDate, @Qty				
				
				INSERT INTO #InvtyTurnOverDtl
				SELECT @TransDate
					 , @TransType
					 , @ReasonCode
					 , @ReasonDesc
					 , @Qty
					 , @Item
					 , @ItemDesc
					 , @ProductCode
					 , @LotNumber
					 , @RefNum
					 , @RefLine
					 , ((matl_unit_cost_php + pi_fg_process_php + pi_resin_php + pi_hidden_profit_php)/ job_qty) * @Qty
					 , ((matl_landed_cost_php) / job_qty) * @Qty

					 --, sf_lbr_cost_php / job_qty
					 --, sf_ovhd_cost_php / job_qty
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
					
				
				INSERT INTO #InvtyTurnOverDtl
				SELECT @TransDate
					 , @TransType
					 , @ReasonCode
					 , @ReasonDesc
					 , @Qty
					 , @Item
					 , @ItemDesc
					 , @ProductCode
					 , @LotNumber
					 , @RefNum
					 , @RefLine
					 , @matl_unit_cost_php * @Qty
					 , 0
			END
		
		END
		ELSE
		BEGIN
		
			EXEC dbo.LSP_ActlCost_GetMatlCostingSp @Item, @LotNumber, @TransDate
					  , @JobQty OUTPUT
					  , @matl_unit_cost_usd OUTPUT, @matl_landed_cost_usd OUTPUT
					  , @pi_fg_process_usd OUTPUT, @pi_resin_usd OUTPUT, @pi_vend_cost_usd OUTPUT, @pi_hidden_profit_usd OUTPUT
					  , @sf_lbr_cost_usd OUTPUT, @sf_ovhd_cost_usd OUTPUT
					  , @fg_lbr_cost_usd OUTPUT, @fg_ovhd_cost_usd OUTPUT
					  , @matl_unit_cost_php OUTPUT, @matl_landed_cost_php OUTPUT
					  , @pi_fg_process_php OUTPUT, @pi_resin_php OUTPUT, @pi_vend_cost_php OUTPUT, @pi_hidden_profit_php OUTPUT
					  , @sf_lbr_cost_php OUTPUT, @sf_ovhd_cost_php OUTPUT
					  , @fg_lbr_cost_php OUTPUT, @fg_ovhd_cost_php OUTPUT
						
			INSERT INTO #InvtyTurnOverDtl
			SELECT @TransDate
					 , @TransType
					 , @ReasonCode
					 , @ReasonDesc
					 , @Qty
					 , @Item
					 , @ItemDesc
					 , @ProductCode
					 , @LotNumber
					 , @RefNum
					 , @RefLine
				 , (ISNULL(@matl_unit_cost_php, 0)  
					 + ISNULL(@pi_fg_process_php, 0)
					 + ISNULL(@pi_resin_php, 0)
					 + ISNULL(@pi_hidden_profit_php, 0)) * @Qty
				 , (ISNULL(@matl_landed_cost_php, 0)) * @Qty
				 
				 --, ISNULL(@sf_lbr_cost_php, 0)
				 --, ISNULL(@sf_ovhd_cost_php, 0)
			
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
	
	
		FETCH NEXT FROM transCrsr INTO
			@TransDate
		  , @TransType
		  , @ReasonCode
		  , @ReasonDesc
		  , @Qty
		  , @Item
		  , @ItemDesc
		  , @ProductCode
		  , @LotNumber
		  , @RefNum
		  , @RefLine	
	
	END	
	
	CLOSE transCrsr
	DEALLOCATE transCrsr

	--***********/
	
	
	DECLARE usageCrsr CURSOR FAST_FORWARD FOR  
	SELECT MAX(TransDate)
		 , SUM(MatlCost)
		 , SUM(LandedCost)
		 , REPLACE(REPLACE(ProductCode, 'SA-',''), 'RM-','')
	FROM #InvtyTurnOverDtl
	GROUP BY YEAR(TransDate), MONTH(TransDate)
		, REPLACE(REPLACE(ProductCode, 'SA-',''), 'RM-','')
	
	OPEN usageCrsr
	FETCH FROM usageCrsr INTO
		@TransDateCrsr  
	  , @MatlUsage  
	  , @LandedUsage  
	  , @ProdCodeCrsr
	  
	WHILE @@FETCH_STATUS = 0
	BEGIN

		EXEC dbo.LSP_NewDM_GetInventorySafetyStockMaterialLandedCostSp @ProdCodeCrsr, @InvtyMatlCost OUTPUT, @InvtyLandedCost OUTPUT, @SafetyMatlCost OUTPUT  
		
		INSERT INTO @DMInvty (  
			trans_date  
		  , usage_matl  
		  , usage_landed  
		  , product_code  
		  , invty_matl_cost  
		  , invty_landed_cost  
		  , safety_matl_cost  
		  , report_group  
		 )  
		SELECT @TransDateCrsr  
			 , @MatlUsage  
			 , @LandedUsage  
			 , @ProdCodeCrsr     
			 , ISNULL(@InvtyMatlCost, 0)  
			 , ISNULL(@InvtyLandedCost, 0)  
			 , ISNULL(@SafetyMatlCost, 0)  
			 , 'USAGE'  

		FETCH NEXT FROM usageCrsr INTO
			@TransDateCrsr  
		  , @MatlUsage  
		  , @LandedUsage  
		  , @ProdCodeCrsr
	
	END
	
	CLOSE usageCrsr
	DEALLOCATE usageCrsr
	
	/****************/
	SET @CurrDate = @StartDate

	WHILE CONVERT(VARCHAR(6), @CurrDate, 112) <= CONVERT(VARCHAR(6), @EndDate, 112) 
	BEGIN

		IF CONVERT(VARCHAR(6), @CurrDate, 112) = CONVERT(VARCHAR(6), @EndDate, 112) 
		BEGIN
			SET @MonthGroup = @MonthGroup + '[' + CONVERT(VARCHAR(6), @CurrDate, 112)  + ']'
			SET @MonthGroup2 = @MonthGroup2 + '[' + CONVERT(VARCHAR(6), @CurrDate, 112)  + '_2]'
			SET @Select1 = @Select1 + 'MAX(ISNULL([' + CONVERT(VARCHAR(6), @CurrDate, 112)  + '],0)) AS _' + CONVERT(VARCHAR(6), @CurrDate, 112) + ''
			SET @Select2 = @Select2 + 'MAX(ISNULL([' + CONVERT(VARCHAR(6), @CurrDate, 112)  + '_2],0)) AS _' + CONVERT(VARCHAR(6), @CurrDate, 112) + '_2'
		END
		ELSE
		BEGIN
			SET @MonthGroup = @MonthGroup + '[' + CONVERT(VARCHAR(6), @CurrDate, 112)  + '], '
			SET @MonthGroup2 = @MonthGroup2 + '[' + CONVERT(VARCHAR(6), @CurrDate, 112)  + '_2], '
			SET @Select1 = @Select1 + 'MAX(ISNULL([' + CONVERT(VARCHAR(6), @CurrDate, 112)  + '],0)) AS _' + CONVERT(VARCHAR(6), @CurrDate, 112) + ', '
			SET @Select2 = @Select2 + 'MAX(ISNULL([' + CONVERT(VARCHAR(6), @CurrDate, 112)  + '_2],0)) AS _' + CONVERT(VARCHAR(6), @CurrDate, 112) + '_2, '
		END

		SET @CurrDate = DATEADD(MONTH, 1,@CurrDate)

	END

	--SELECT @MonthGroup, @MonthGroup2, @Select1, @Select2

	SELECT @SqlStr = 'SELECT product_code ' 
					+ ', ' + @Select1 
					+ ', ' + @Select2 
					+ ', invty_matl_cost  
					   , invty_landed_cost  
					   , safety_matl_cost  
					   , report_group'
					+ ' FROM (  SELECT CONVERT(VARCHAR(6), trans_date, 112) AS [YYYYMM_1]--, * 
						 , usage_matl
						 , (CONVERT(VARCHAR(6), trans_date, 112) + ''_2'') AS [YYYYMM_2]
						 , usage_landed
						 , product_code
						 , invty_matl_cost  
						 , invty_landed_cost  
						 , safety_matl_cost  
						 , report_group
					FROM @DMInvtyTbl
					WHERE report_group = ''USAGE''
				  ) rpt
			PIVOT (SUM(usage_matl) FOR [YYYYMM_1] IN ('+ @MonthGroup +') ) pvt
			PIVOT (SUM(usage_landed) FOR [YYYYMM_2] IN ('+ @MonthGroup2 +') ) pvt2
			GROUP BY product_code, invty_matl_cost, invty_landed_cost, safety_matl_cost, report_group;'
			
	SET @ParmDefinition = N'@DMInvtyTbl DMInvtyTurnOverType READONLY';  
	
	INSERT INTO @report_set (
		product_code
	  , usage_M1
	  , usage_M2
	  , usage_M3
	  , usage_M4
	  , usage_M5
	  , usage_M6
	  , usage_M7
	  , usage_M8
	  , usage_M9
	  , usage_M10
	  , usage_M11
	  , usage_M12
	  , usage_L1
	  , usage_L2
	  , usage_L3
	  , usage_L4
	  , usage_L5
	  , usage_L6
	  , usage_L7
	  , usage_L8
	  , usage_L9
	  , usage_L10
	  , usage_L11
	  , usage_L12
	  , invty_matl_cost
	  , invty_landed_cost
	  , safety_matl_cost
	  , report_group
	)
	EXEC sp_executesql @SqlStr, @ParmDefinition, @DMInvtyTbl = @DMInvty
	
	/********/
	
	DECLARE ProdCodeCrsr CURSOR FAST_FORWARD FOR  
	SELECT product_code  
	FROM @report_set  
	GROUP BY product_code  
	  
	OPEN ProdCodeCrsr  
	FETCH FROM ProdCodeCrsr INTO  
	 @ProdCode
	   
	WHILE (@@FETCH_STATUS = 0)  
	BEGIN  
	  
	 UPDATE @report_set  
	 SET M1 = (SELECT ISNULL(SUM(usage_matl),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(@StartDate)  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 1, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 2, @StartDate))) )  
	   , M2 = (SELECT ISNULL(SUM(usage_matl),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 1, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 2, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 3, @StartDate))) )  
	   , M3 = (SELECT ISNULL(SUM(usage_matl),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 2, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 3, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 4, @StartDate))) )  
	   , M4 = (SELECT ISNULL(SUM(usage_matl),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 3, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 4, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 5, @StartDate))) )  
	   , M5 = (SELECT ISNULL(SUM(usage_matl),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 4, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 5, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 6, @StartDate))) )  
	   , M6 = (SELECT ISNULL(SUM(usage_matl),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 5, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 6, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 7, @StartDate))) )  
	   , M7 = (SELECT ISNULL(SUM(usage_matl),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 6, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 7, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 8, @StartDate))) )  
	   , M8 = (SELECT ISNULL(SUM(usage_matl),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 7, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 8, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 9, @StartDate))) )  
	   , M9 = (SELECT ISNULL(SUM(usage_matl),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 8, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 9, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 10, @StartDate))) )  
	   , M10 = (SELECT ISNULL(SUM(usage_matl),0)  
				  FROM @DMInvty  
				  WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 9, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 10, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 11, @StartDate))) )  
	 --LANDED COST USAGE  
	   , L1 = (SELECT ISNULL(SUM(usage_landed),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(@StartDate)  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 1, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 2, @StartDate))) )  
	   , L2 = (SELECT ISNULL(SUM(usage_landed),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 1, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 2, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 3, @StartDate))) )  
	   , L3 = (SELECT ISNULL(SUM(usage_landed),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 2, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 3, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 4, @StartDate))) )  
	   , L4 = (SELECT ISNULL(SUM(usage_landed),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 3, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 4, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 5, @StartDate))) )  
	   , L5 = (SELECT ISNULL(SUM(usage_landed),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 4, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 5, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 6, @StartDate))) )  
	   , L6 = (SELECT ISNULL(SUM(usage_landed),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 5, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 6, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 7, @StartDate))) )  
	   , L7 = (SELECT ISNULL(SUM(usage_landed),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 6, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 7, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 8, @StartDate))) )  
	   , L8 = (SELECT ISNULL(SUM(usage_landed),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 7, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 8, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 9, @StartDate))) )  
	   , L9 = (SELECT ISNULL(SUM(usage_landed),0)  
				 FROM @DMInvty  
				 WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 8, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 9, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 10, @StartDate))) )  
	   , L10 = (SELECT ISNULL(SUM(usage_landed),0)  
				  FROM @DMInvty  
				  WHERE product_code = @ProdCode  
				AND (MONTH(trans_date) = MONTH(DATEADD(MONTH, 9, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 10, @StartDate))  
				 OR MONTH(trans_date) = MONTH(DATEADD(MONTH, 11, @StartDate))) )  
	 WHERE product_code = @ProdCode  
	   
	 UPDATE @report_set  
	 SET MAX_3months = (SELECT TOP(1)(SELECT MAX(m)   
			  FROM (VALUES (M1), (M2), (M3), (M4), (M5), (M6), (M7), (M8), (M9), (M10)) AS value(m))  
			FROM @DMInvty  
			WHERE product_code = @ProdCode)  
	   , L_MAX_3Months = (SELECT TOP(1)(SELECT MAX(l)   
			  FROM (VALUES (L1), (L2), (L3), (L4), (L5), (L6), (L7), (L8), (L9), (L10)) AS value(l))  
			FROM @DMInvty  
			WHERE product_code = @ProdCode)  
	 WHERE product_code = @ProdCode  
	  
	 FETCH NEXT FROM ProdCodeCrsr INTO  
		@ProdCode  
	  
	END  
	  
	CLOSE ProdCodeCrsr  
	DEALLOCATE ProdCodeCrsr 

	
	IF @IsShowDetail = 1
	BEGIN
	
		INSERT INTO @report_set (
			trans_date  
		  , trans_type  
		  , reason_code  
		  , reason_desc  
		  , qty  
		  , usage_matl  
		  , usage_landed  
		  , item  
		  , item_desc  
		  , product_code  
		  , lot  
		  , ref_num  
		  , ref_line  
		  , report_group  
		 )  
		 SELECT TransDate  
		   , TransType  
		   , ReasonCode  
		   , ReasonDesc  
		   , Qty  
		   , MatlCost  
		   , LandedCost  
		   , Item  
		   , ItemDesc  
		   , ProductCode  
		   , LotNumber  
		   , RefNum  
		   , RefLine  
		   , 'DETAILED'  
		 FROM #InvtyTurnOverDtl	
	END
	
	
	SELECT * 
	--INTO Rpt_InvtyTurnover
	FROM @report_set  
	ORDER BY report_group DESC, product_code, trans_date  
	
END