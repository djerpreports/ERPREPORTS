--ALTER PROCEDURE LSP_Rpt_NewDM_RMBreakdownPerJOSp (
DECLARE
	@JobOrder				NVARCHAR(20) = '20-0000558'--'19-0002507'
  , @PONumber				NVARCHAR(20) = ''
  , @Quantity				INT			 = NULL
--) AS
BEGIN

	IF OBJECT_ID('tempdb..#BOMStdCost') IS NOT NULL
		DROP TABLE #BOMStdCost	
	IF OBJECT_ID('tempdb..#DMActualCost') IS NOT NULL
		DROP TABLE #DMActualCost
	IF OBJECT_ID('tempdb..#RMBreakdownActualCost') IS NOT NULL
		DROP TABLE #RMBreakdownActualCost
		
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
	
	CREATE TABLE #RMBreakdownActualCost (
		item						NVARCHAR(60)
	  , [Level]						INT
	  , Parent						NVARCHAR(20)
	  , oper_num					INT
	  , sequence					NVARCHAR(3)
	  , subsequence					NVARCHAR(50)
	  , matl						NVARCHAR(60)
	  , matl_qty					DECIMAL(18,8)
	  , lot_no						NVARCHAR(50)
	  , trans_date					DATETIME
	  , job_qty						BIGINT
	  , job_matl_qty				DECIMAL(18,8)
	  , actl_matl_qty				DECIMAL(18,8)
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

	DECLARE
		@JobSuffix			SuffixType
	  , @JobQtyRelease		INT
	  , @JobDate			DateType
	  , @JobItem			ItemType
	  , @CurrLevel			INT
	  , @MaxLevel			INT
	  , @FGStdLbrHrs		DECIMAL(18,10)
	  , @FGActLbrHrs		DECIMAL(18,10)

	SELECT TOP(1) 
		   @JobOrder = ISNULL(NULLIF(@JobOrder,''), job)
		 , @PONumber = ISNULL(NULLIF(@PONumber,''), Uf_ponum)
		 , @JobSuffix = suffix
		 , @JobDate = job_date
		 , @JobItem = item
		 , @JobQtyRelease = ISNULL(@Quantity, qty_complete) --qty_released
		 
	FROM job
	WHERE job = @JobOrder
	   OR Uf_ponum = @PONumber

	SELECT @FGStdLbrHrs = SUM(js.run_lbr_hrs * jt.qty_complete)
		 , @FGActLbrHrs = SUM(ISNULL(jt.a_hrs,0))
	FROM jobtran AS jt 
			JOIN job AS j 
				ON jt.job = j.job 
				  AND jt.suffix = j.suffix 
			JOIN  item AS i 
				ON j.item = i.item 
			JOIN  jrt_sch AS js 
				ON i.job = js.job 
				  AND i.suffix = js.suffix 
				  AND jt.oper_num = js.oper_num
	WHERE j.job = @JobOrder
		   
	--SELECT @JobOrder, @PONumber
	--	 , @JobSuffix
	--	 , @JobDate
	--	 , @JobItem
	--	 , @JobQtyRelease

	INSERT INTO #BOMStdCost
	EXEC dbo.LSP_DM_StdCost_GetCurrentMatlCostingSp @JobItem, @JobDate
	
	--SELECT @JobOrder, @JobSuffix, @JobItem, @JobDate, @JobQtyRelease

	INSERT INTO #DMActualCost
	EXEC dbo.LSP_DM_ActlCost_GetJobMatlTransCostingSp @JobOrder, @JobSuffix, @JobItem, @JobDate, @JobQtyRelease
	
	SELECT @MaxLevel = MAX([Level])
		 , @CurrLevel = 2
	FROM #DMActualCost

	INSERT INTO #RMBreakdownActualCost
	SELECT item
	  , [Level]
	  , Parent
	  , oper_num
	  , sequence
	  , subsequence
	  , matl
	  , matl_qty
	  , lot_no
	  , trans_date
	  , job_qty
	  , job_qty		
	  , matl_qty
	  , (matl_unit_cost_php / matl_qty)
	  , (matl_landed_cost_php / matl_qty)
	  , (pi_fg_process_php / matl_qty)
	  , (pi_resin_php / matl_qty)
	  , (pi_vend_cost_php / matl_qty)
	  , (pi_hidden_profit_php / matl_qty)
	  , (sf_lbr_cost_php / matl_qty)
	  , (sf_ovhd_cost_php / matl_qty)
	  , (fg_lbr_cost_php / matl_qty)
	  , (fg_ovhd_cost_php / matl_qty)
	 
	FROM #DMActualCost
	WHERE [Level] = 0
	
	INSERT INTO #RMBreakdownActualCost
	SELECT a.item
	  , a.[Level]
	  , a.Parent
	  , a.oper_num
	  , a.sequence
	  , a.subsequence
	  , a.matl
	  , a.matl_qty
	  , a.lot_no
	  , a.trans_date
	  , a.job_qty
	  , (a.matl_qty / a1.job_qty)
	  , ((a.matl_qty / a1.job_qty) * a1.job_matl_qty)
	  , a.matl_unit_cost_php / a.matl_qty
	  , a.matl_landed_cost_php / a.matl_qty
	  , a.pi_fg_process_php / a.matl_qty
	  , a.pi_resin_php / a.matl_qty
	  , a.pi_vend_cost_php / a.matl_qty
	  , a.pi_hidden_profit_php / a.matl_qty
	  , a.sf_lbr_cost_php / a.matl_qty
	  , a.sf_ovhd_cost_php / a.matl_qty
	  , a.fg_lbr_cost_php / a.matl_qty
	  , a.fg_ovhd_cost_php / a.matl_qty
	 
	FROM #DMActualCost AS a
		LEFT OUTER JOIN #RMBreakdownActualCost AS a1
			ON a1.[Level] = 0
	WHERE a.[Level] = 1	

	WHILE(@CurrLevel <= @MaxLevel)
	BEGIN
	--SELECT * FROM #RMBreakdownActualCost
		INSERT INTO #RMBreakdownActualCost --(
				--item
			 -- , [Level]
			 -- , Parent
			 -- , oper_num
			 -- , sequence
			 -- , subsequence
			 -- , matl
			 -- , matl_qty
			 -- , lot_no
			 -- , trans_date
			 -- , job_qty
			 -- , matl_unit_cost_php
			 -- , matl_landed_cost_php
			 -- , pi_fg_process_php
			 -- , pi_resin_php
			 -- , pi_vend_cost_php
			 -- , pi_hidden_profit_php
			 -- , sf_lbr_cost_php
			 -- , sf_ovhd_cost_php
			 -- , fg_lbr_cost_php
			 -- , fg_ovhd_cost_php )
		SELECT a.item
		     , a.[Level]
		     , a.Parent
		     , a.oper_num
		     , a.sequence
		     , a.subsequence
		     , a.matl
		     , a.matl_qty
		     , a.lot_no
		     , a.trans_date
		     , a.job_qty
		     --, a.matl_qty, a1.job_qty, a1.job_matl_qty
		     , CAST(ISNULL((a.matl_qty / NULLIF(a1.job_qty,0)),0) AS DECIMAL(18,8))
		     , CAST(ISNULL((a.matl_qty / NULLIF(a1.job_qty,0)),0) AS DECIMAL(18,8)) * a1.actl_matl_qty
		     --, ((a.matl_qty / a1.job_qty) * a1.job_matl_qty)
		     -- CASE WHEN a.[Level] = 2 THEN a1.job_matl_qty ELSE a1.actl_matl_qty END
		  --   ,  (matl_qty 
				--   / (SELECT actl_matl_qty 
				--		FROM #RMBreakdownActualCost AS a1 
				--		WHERE (CAST(a1.[Level] AS NVARCHAR(2)) + '.' + CAST(a1.[oper_num] AS NVARCHAR(2)) + '.' + a1.sequence) = a.Parent 
				--			  AND a1.[Level] = (a.[Level] -1)
				--			  AND a1.matl = a.item 
				--			  AND 1 = CASE WHEN CAST(a.sequence AS INT) < 10 AND a1.subsequence = LEFT(a.subsequence,LEN(a.subsequence) - 2) 
				--								THEN 1
				--						   WHEN CAST(a.sequence AS INT) >= 10 AND a1.subsequence = LEFT(a.subsequence,LEN(a.subsequence) - 3)
				--								THEN 1
				--						   ELSE 0 
				--					  END))
				--* (SELECT matl_qty 
				--		FROM #DMActualCost AS a2
				--		WHERE (CAST(a2.[Level] AS NVARCHAR(2)) + '.' + CAST(a2.[oper_num] AS NVARCHAR(2)) + '.' + a2.sequence) = a.Parent 
				--			  AND a2.[Level] = (a.[Level] -1)
				--			  AND a2.matl = a.item 
				--			  AND 1 = CASE WHEN CAST(a.sequence AS INT) < 10 AND a2.subsequence = LEFT(a.subsequence,LEN(a.subsequence) - 2) 
				--								THEN 1
				--						   WHEN CAST(a.sequence AS INT) >= 10 AND a2.subsequence = LEFT(a.subsequence,LEN(a.subsequence) - 3)
				--								THEN 1
				--						   ELSE 0 
				--					  END)
		     , ISNULL((a.matl_unit_cost_php / NULLIF(a.matl_qty,0)),0) --* ((a.matl_qty / a1.job_qty) * a1.job_matl_qty)
		     , ISNULL((a.matl_landed_cost_php / NULLIF(a.matl_qty,0)), 0) --* ((a.matl_qty / a1.job_qty) * a1.actl_matl_qty)
		     , ISNULL((a.pi_fg_process_php / NULLIF(a.matl_qty,0)), 0)  --* ((a.matl_qty / a1.job_qty) * a1.actl_matl_qty)
		     , ISNULL((a.pi_resin_php / NULLIF(a.matl_qty,0)), 0)  --* ((a.matl_qty / a1.job_qty) * a1.actl_matl_qty)
		     , ISNULL((a.pi_vend_cost_php / NULLIF(a.matl_qty,0)), 0)  --* ((a.matl_qty / a1.job_qty) * a1.actl_matl_qty)
		     , ISNULL((a.pi_hidden_profit_php / NULLIF(a.matl_qty,0)), 0)  --* ((a.matl_qty / a1.job_qty) * a1.actl_matl_qty)
		     , ISNULL((a.sf_lbr_cost_php / NULLIF(a.matl_qty,0)), 0)  --* ((a.matl_qty / a1.job_qty) * a1.actl_matl_qty)
		     , ISNULL((a.sf_ovhd_cost_php / NULLIF(a.matl_qty,0)), 0) --* ((a.matl_qty / a1.job_qty) * a1.actl_matl_qty)
		     , ISNULL((a.fg_lbr_cost_php / NULLIF(a.matl_qty,0)), 0) --* ((a.matl_qty / a1.job_qty) * a1.actl_matl_qty)
		     , ISNULL((a.fg_ovhd_cost_php / NULLIF(a.matl_qty,0)), 0) --* ((a.matl_qty / a1.job_qty) * a1.actl_matl_qty)
	
		FROM #DMActualCost AS a
			LEFT OUTER JOIN #RMBreakdownActualCost AS a1 
						ON (CAST(a1.[Level] AS NVARCHAR(2)) + '.' + CAST(a1.[oper_num] AS NVARCHAR(2)) + '.' + a1.sequence) = a.Parent 
							  AND a1.[Level] = (a.[Level] -1)
							  AND a1.matl = a.item 
							  AND 1 = CASE WHEN CAST(a.sequence AS INT) < 10 AND a1.subsequence = LEFT(a.subsequence,LEN(a.subsequence) - 2) 
												THEN 1
										   WHEN CAST(a.sequence AS INT) >= 10 AND a1.subsequence = LEFT(a.subsequence,LEN(a.subsequence) - 3)
												THEN 1
										   ELSE 0 
									  END
		WHERE a.[Level] = @CurrLevel
		
		--IF @CurrLevel = 2
		--SELECT 	item, matl, matl_unit_cost_php, matl_qty
		--FROM #DMActualCost	
		--WHERE [Level] = @CurrLevel
	
		SET @CurrLevel = @CurrLevel + 1
	
	END

	--SELECT * FROM #BOMStdCost
	--ORDER BY subsequence, CAST(sequence AS INT), [Level]

	--SELECT * FROM #DMActualCost
	--SELECT * FROM #RMBreakdownActualCost
	--ORDER BY subsequence, CAST(sequence AS INT), [Level]
	
	SELECT @JobOrder	AS JONum
		 , @PONumber	AS PONum
		 , ac.matl
		 , i.description AS matl_desc
		 , CASE WHEN ac.[Level] = 0
					THEN @FGStdLbrHrs
				ELSE 0 END
			 AS StdLbrHrs
		 , CASE WHEN ac.[Level] = 0
					THEN @FGActLbrHrs
				ELSE 0 END
			 AS ActlLbrHrs
		 , ISNULL(dbo.LSP_fn_GetCurrencyConversion(ac.trans_date,'USD','PHP', std.matl_unit_cost), 0) AS std_matl_unit
		 , ISNULL(dbo.LSP_fn_GetCurrencyConversion(ac.trans_date,'USD','PHP', std.pi_process_cost), 0) AS std_process_unit
		 , ISNULL(dbo.LSP_fn_GetCurrencyConversion(ac.trans_date,'USD','PHP', std.pi_resin_cost), 0) AS pi_resin_unit
		 , ISNULL(dbo.LSP_fn_GetCurrencyConversion(ac.trans_date,'USD','PHP', std.pi_hidden_profit), 0) AS pi_hidden_unit
		 , ISNULL(dbo.LSP_fn_GetCurrencyConversion(ac.trans_date,'USD','PHP', std.sf_labr_cost), 0) AS sf_lbr_unit
		 , ISNULL(dbo.LSP_fn_GetCurrencyConversion(ac.trans_date,'USD','PHP', std.sf_ovhd_cost), 0) AS sf_ovhd_unit
		 , ISNULL(dbo.LSP_fn_GetCurrencyConversion(ac.trans_date,'USD','PHP', std.fg_labr_cost), 0) AS fg_lbr_unit
		 , ISNULL(dbo.LSP_fn_GetCurrencyConversion(ac.trans_date,'USD','PHP', std.fg_ovhd_cost), 0) AS fg_ovhd_unit
		 , ( ISNULL(dbo.LSP_fn_GetCurrencyConversion(ac.trans_date,'USD','PHP', std.matl_unit_cost), 0)
			 + ISNULL(dbo.LSP_fn_GetCurrencyConversion(ac.trans_date,'USD','PHP', std.pi_process_cost), 0)
			 + ISNULL(dbo.LSP_fn_GetCurrencyConversion(ac.trans_date,'USD','PHP', std.pi_resin_cost), 0)
			 + ISNULL(dbo.LSP_fn_GetCurrencyConversion(ac.trans_date,'USD','PHP', std.pi_hidden_profit), 0)
			 + ISNULL(dbo.LSP_fn_GetCurrencyConversion(ac.trans_date,'USD','PHP', std.sf_labr_cost), 0)
			 + ISNULL(dbo.LSP_fn_GetCurrencyConversion(ac.trans_date,'USD','PHP', std.sf_ovhd_cost), 0)
			 + ISNULL(dbo.LSP_fn_GetCurrencyConversion(ac.trans_date,'USD','PHP', std.fg_labr_cost), 0)
			 + ISNULL(dbo.LSP_fn_GetCurrencyConversion(ac.trans_date,'USD','PHP', std.fg_ovhd_cost), 0) 
		  ) AS total_std_unit
		--/**********
		 , ISNULL(ac.[Level], 0) AS [Level]
		 , ISNULL(ac.sequence, 0) AS sequence
		 , ISNULL(ac.subsequence, 0) AS subsequence
		 , ISNULL(ac.lot_no, 0) AS lot_no
		 , ISNULL(ac.matl_qty, 0) AS matl_qty
		 , ISNULL(ac.job_qty, 0) AS job_qty
		 , ISNULL(ac.job_matl_qty, 0) AS job_matl_qty
		 , ISNULL(ac.actl_matl_qty, 0) AS actl_matl_qty
		 , ISNULL(ac.matl_unit_cost_php, 0) AS matl_unit_cost_php
		 , ISNULL(ac.matl_landed_cost_php, 0) AS matl_landed_cost_php
		 , ISNULL(ac.pi_fg_process_php, 0) AS pi_fg_process_php
		 , ISNULL(ac.pi_resin_php, 0) AS pi_resin_php
		 , ISNULL(ac.pi_hidden_profit_php, 0) AS pi_hidden_profit_php
		 , ISNULL(ac.sf_lbr_cost_php, 0) AS sf_lbr_cost_php
		 , ISNULL(ac.sf_ovhd_cost_php, 0) AS sf_ovhd_cost_php
		 , ISNULL(ac.fg_lbr_cost_php, 0) AS fg_lbr_cost_php
		 , ISNULL(ac.fg_ovhd_cost_php, 0) AS fg_ovhd_cost_php
		 , (ISNULL(ac.matl_unit_cost_php, 0)
			 + ISNULL(ac.matl_landed_cost_php, 0)
			 + ISNULL(ac.pi_fg_process_php, 0)
			 + ISNULL(ac.pi_resin_php, 0)
			 + ISNULL(ac.pi_hidden_profit_php, 0)
			 + ISNULL(ac.sf_lbr_cost_php, 0)
			 + ISNULL(ac.sf_ovhd_cost_php, 0)
			 + ISNULL(ac.fg_lbr_cost_php, 0)
			 + ISNULL(ac.fg_ovhd_cost_php, 0)
			) as total_actl_unit
		  , (ISNULL(ac.matl_unit_cost_php, 0)
			 + ISNULL(ac.pi_fg_process_php, 0)
			 + ISNULL(ac.pi_resin_php, 0)
			 + ISNULL(ac.pi_hidden_profit_php, 0)
			 + ISNULL(ac.sf_lbr_cost_php, 0)
			 + ISNULL(ac.sf_ovhd_cost_php, 0)
			 + ISNULL(ac.fg_lbr_cost_php, 0)
			 + ISNULL(ac.fg_ovhd_cost_php, 0)
			) as nolanded_actl_unit
		--**********/
		 --, CASE WHEN [Level] = 0 OR [Level] > 1
			--		THEN (matl_unit_cost_php * actl_matl_qty)				
			--	ELSE (matl_unit_cost_php / actl_matl_qty) * job_matl_qty END
	  --   , CASE WHEN [Level] > 1 
			--		THEN (matl_landed_cost_php * actl_matl_qty)
			--	ELSE (matl_landed_cost_php / actl_matl_qty) * job_matl_qty END
	  --   , CASE WHEN [Level] > 1 
			--		THEN ( * actl_matl_qty)
			--	ELSE (pi_fg_process_php / actl_matl_qty) * job_matl_qty END
	  --   , CASE WHEN [Level] > 1 
			--		THEN (pi_resin_php * actl_matl_qty)
			--	ELSE (pi_resin_php / actl_matl_qty) * job_matl_qty END
	  --   , CASE WHEN [Level] > 1 
			--		THEN (pi_vend_cost_php * actl_matl_qty)
			--	ELSE (pi_vend_cost_php / actl_matl_qty) * job_matl_qty END
	  --   , CASE WHEN [Level] > 1 
			--		THEN (pi_hidden_profit_php * actl_matl_qty)
			--	ELSE (pi_hidden_profit_php / actl_matl_qty) * job_matl_qty END
	  --   , CASE WHEN [Level] > 1 
			--		THEN (sf_lbr_cost_php * actl_matl_qty)
			--	ELSE (sf_lbr_cost_php / actl_matl_qty) * job_matl_qty END
	  --   , CASE WHEN [Level] > 1 
			--		THEN (sf_ovhd_cost_php * actl_matl_qty)
			--	ELSE (sf_ovhd_cost_php / actl_matl_qty) * job_matl_qty END
	  --   , CASE WHEN [Level] > 1 
			--		THEN (fg_lbr_cost_php * actl_matl_qty)
			--	ELSE (fg_lbr_cost_php / actl_matl_qty) * job_matl_qty END
	  --   , CASE WHEN [Level] > 1 
			--		THEN (fg_ovhd_cost_php * actl_matl_qty)
			--	ELSE (fg_ovhd_cost_php / actl_matl_qty) * job_matl_qty END
	
	FROM #RMBreakdownActualCost AS ac
		 JOIN item AS i 
			ON ac.matl = i.item
		 LEFT OUTER JOIN #BOMStdCost AS std
			ON ac.matl = std.matl
			  AND ac.[Level] = std.[Level]
			  AND ac.oper_num = std.oper_num
			  --AND ac.sequence = std.sequence
	ORDER BY ac.subsequence, CAST(ac.sequence AS INT), ac.[Level]

END