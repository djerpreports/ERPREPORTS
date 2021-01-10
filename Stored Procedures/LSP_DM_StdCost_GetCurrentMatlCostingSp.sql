--LSP_DM_StdCost_GetCurrentMatlCostingSp 'FG-DK-100D'

--ALTER PROCEDURE LSP_DM_StdCost_GetCurrentMatlCostingSp (
DECLARE
	@Item				ItemType = 'FG-SR4-017'
  , @TransDate			DateType = '05/29/2020'
--) AS 
BEGIN

	IF OBJECT_ID('tempdb..#itemMatl') IS NOT NULL
		DROP TABLE #itemMatl
	IF OBJECT_ID('tempdb..#subMatl') IS NOT NULL
		DROP TABLE #subMatl
	IF OBJECT_ID('tempdb..#SFMatl') IS NOT NULL
		DROP TABLE #SFMatl
	IF OBJECT_ID('tempdb..#BOMCost') IS NOT NULL
		DROP TABLE #BOMCost
		
	CREATE TABLE #itemMatl (
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
		
	DECLARE
		@Matl				ItemType
	  , @Level				INT
	  , @OperNum			OperNumType
	  , @Sequence			SequenceType
	  , @SubSequence		NVARCHAR(50)
	  , @CurrLevel			INT				= 1
	  , @MatlUnitCost		DECIMAL(18,10)
	  , @PIProcessCost		DECIMAL(18,10)
	  , @PIResinCost		DECIMAL(18,10)
	  , @PIHiddenProfit		DECIMAL(18,10)
	  , @SFLabrCost			DECIMAL(18,10)
	  , @SFOvhdCost			DECIMAL(18,10)
	  , @FGLabrCost			DECIMAL(18,10)
	  , @FGOvhdCost			DECIMAL(18,10)

	  , @Ctr				INT
	  , @LevelCtr			INT	  

	INSERT INTO #itemMatl 
		(	item
		  , [Level]
		  , Parent
		  , oper_num
		  , sequence
		  , subsequence
		  , matl
		  , matl_qty
		)
	SELECT i.item
		 --, i.description
		 --, i.job
		 --, i.suffix
		 , @CurrLevel AS [Level]
		 , CAST(0 AS NVARCHAR(20)) AS Parent
		 , jm.oper_num
		 , jm.sequence
		 , CAST(jm.oper_num AS NVARCHAR(5)) + '_' + CAST(jm.sequence AS NVARCHAR(50)) AS subsequence
		 , jm.item AS matl
		 , jm.matl_qty
	FROM item AS i
		JOIN jobmatl AS jm
			ON i.job = jm.job AND i.suffix = jm.suffix
	WHERE i.item = @Item
	  AND i.stat IN ('A', 'S')
	
	UNION
	SELECT @Item
		 , 0
		 , '0'
		 , '0'
		 , 0
		 , '0'
		 , @Item
		 , 1
	ORDER BY jm.oper_num, jm.sequence
	
	SELECT item
		 , [Level]
		 , Parent
		 , oper_num
		 , sequence
		 , subsequence
		 , matl
		 , matl_qty
	INTO #SFMatl
	FROM #itemMatl
	WHERE (matl LIKE 'SF-%' OR matl LIKE 'FG-%')
	  AND [Level] > 0
	
	SELECT item
		 , [Level]
		 , Parent
		 , oper_num
		 , sequence
		 , subsequence
		 , matl
		 , matl_qty
	INTO #subMatl
	FROM #itemMatl
	WHERE (matl LIKE 'SF-%' OR matl LIKE 'FG-%')
	  AND [Level] > 0
	
	WHILE (SELECT COUNT(*) FROM #SFMatl) > 0
	BEGIN
		
		SELECT @CurrLevel = @CurrLevel + 1
		TRUNCATE TABLE #subMatl

		DECLARE matlCrsr CURSOR FAST_FORWARD FOR	
		SELECT matl
			 , [Level]
			 , oper_num
			 , sequence
			 , subsequence
		FROM #SFMatl
		WHERE matl LIKE 'SF-%'	
		
		OPEN matlCrsr
		FETCH FROM matlCrsr INTO
			@Matl
		  , @Level
		  , @OperNum
		  , @Sequence
		  , @SubSequence
		
		WHILE (@@FETCH_STATUS = 0)
		BEGIN					
		
			INSERT INTO #itemMatl
					(	item
					  , [Level]
					  , Parent
					  , oper_num
					  , sequence
					  , subsequence
					  , matl
					  , matl_qty
					)
			SELECT i.item
				 , @CurrLevel AS [Level]
				 , CAST(@Level AS NVARCHAR(5)) + '.' + CAST(@OperNum AS NVARCHAR(5)) + '.' + CAST(@Sequence AS NVARCHAR(5))AS Parent
				 , jm.oper_num
				 , jm.sequence
				 , @SubSequence + '.' + CAST(jm.sequence AS NVARCHAR(10))
				 , jm.item AS matl
				 , jm.matl_qty
			FROM item AS i
				JOIN jobmatl AS jm
					ON i.job = jm.job AND i.suffix = jm.suffix
			WHERE i.item = @Matl
			  AND i.stat IN ('A', 'S')
			ORDER BY jm.oper_num, jm.sequence
			
			INSERT INTO #subMatl
			SELECT i.item
				 , @Level + 1 AS [Level]
				 , CAST(@Level AS NVARCHAR(5)) + '.' + CAST(@OperNum AS NVARCHAR(5)) + '.' + CAST(@Sequence AS NVARCHAR(5))AS Parent
				 , jm.oper_num
				 , jm.sequence
				 , @SubSequence + '.' + CAST(jm.sequence AS NVARCHAR(10))
				 , jm.item AS matl
				 , jm.matl_qty
			FROM item AS i
				JOIN jobmatl AS jm
					ON i.job = jm.job AND i.suffix = jm.suffix
			WHERE i.item = @Matl
			  AND i.stat IN ('A', 'S')
			  AND jm.item LIKE 'SF-%'
			ORDER BY jm.oper_num, jm.sequence
		
			FETCH NEXT FROM matlCrsr INTO
				@Matl
			  , @Level
			  , @OperNum
			  , @Sequence
			  , @SubSequence
		
		END
		
		CLOSE matlCrsr
		DEALLOCATE matlCrsr
		
		TRUNCATE TABLE #SFMatl
		INSERT INTO #SFMatl
		SELECT * FROM #subMatl	
	END


	DECLARE matlCrsr CURSOR FAST_FORWARD FOR
	SELECT DISTINCT matl
	FROM #itemMatl
	
	OPEN matlCrsr
	FETCH FROM matlCrsr INTO
		@Matl
		
	WHILE (@@FETCH_STATUS = 0)
	BEGIN
		SELECT @MatlUnitCost = 0
			 , @PIProcessCost = 0
			 , @PIResinCost  = 0
			 , @PIHiddenProfit  = 0
			 , @SFLabrCost	 = 0
			 , @SFOvhdCost	 = 0
			 , @FGLabrCost	 = 0
			 , @FGOvhdCost = 0
	
		EXEC dbo.LSP_StdCost_GetMatlCostingSp 
					@Matl, @TransDate, @MatlUnitCost OUTPUT
					, @PIProcessCost	OUTPUT, @PIResinCost OUTPUT, @PIHiddenProfit OUTPUT
					, @SFLabrCost	OUTPUT, @SFOvhdCost	OUTPUT
					, @FGLabrCost	OUTPUT, @FGOvhdCost OUTPUT
					
		UPDATE #itemMatl
		SET matl_unit_cost = ISNULL(@MatlUnitCost, 0)
		  , pi_process_cost = ISNULL(@PIProcessCost, 0)
		  , pi_resin_cost = ISNULL(@PIResinCost, 0)
		  , pi_hidden_profit = ISNULL(@PIHiddenProfit, 0)
		  , sf_labr_cost = ISNULL(@SFLabrCost, 0)
		  , sf_ovhd_cost = ISNULL(@SFOvhdCost, 0)
		  , fg_labr_cost = ISNULL(@FGLabrCost, 0)
		  , fg_ovhd_cost = ISNULL(@FGOvhdCost, 0)
		WHERE matl = @Matl
	
		FETCH NEXT FROM matlCrsr INTO
			@Matl
	
	END
	
	CLOSE matlCrsr
	DEALLOCATE matlCrsr
	

	--SELECT * 
	--FROM #itemMatl
	--ORDER BY subsequence, sequence, [Level]

	SELECT @LevelCtr = MAX([Level])
	FROM #itemMatl
	
	SELECT * 
	INTO #BOMCost
	FROM #itemMatl
	WHERE [Level] = @LevelCtr
	
	WHILE @LevelCtr >= 0
	BEGIN		
		
		--SELECT Parent, CAST( REVERSE(SUBSTRING(REVERSE(subsequence), CHARINDEX('.', REVERSE(subsequence))+1, 100)) AS NVARcHAR(50))AS subs_parent
		--	, SUM(matl_qty * matl_unit_cost) AS child_matl
		--	, SUM(matl_qty * pi_process_cost) AS child_pi_process
		--	, SUM(matl_qty * pi_resin_cost) AS child_pi_resin
		--	, SUM(matl_qty * pi_hidden_profit) AS child_pi_hidden
		--	, SUM(matl_qty * sf_labr_cost)  AS child_sf_lbr
		--	, SUM(matl_qty * sf_ovhd_cost)  AS child_sf_ovhd
		--	, SUM(matl_qty * fg_labr_cost)  AS child_fg_lbr
		--	, SUM(matl_qty * fg_ovhd_cost)  AS child_fg_ovhd
		--FROM #BOMCost
		--WHERE [Level] = @LevelCtr
		--GROUP BY Parent, REVERSE(SUBSTRING(REVERSE(subsequence), CHARINDEX('.', REVERSE(subsequence))+1, 100));
		
		WITH childCost AS (
		SELECT Parent, CAST( REVERSE(SUBSTRING(REVERSE(subsequence), CHARINDEX('.', REVERSE(subsequence))+1, 100)) AS NVARcHAR(50))AS subs_parent
			, SUM(matl_qty * matl_unit_cost) AS child_matl
			, SUM(matl_qty * pi_process_cost) AS child_pi_process
			, SUM(matl_qty * pi_resin_cost) AS child_pi_resin
			, SUM(matl_qty * pi_hidden_profit) AS child_pi_hidden
			, SUM(matl_qty * sf_labr_cost)  AS child_sf_lbr
			, SUM(matl_qty * sf_ovhd_cost)  AS child_sf_ovhd
			, SUM(matl_qty * fg_labr_cost)  AS child_fg_lbr
			, SUM(matl_qty * fg_ovhd_cost)  AS child_fg_ovhd
		FROM #BOMCost
		WHERE [Level] = @LevelCtr
		GROUP BY Parent, REVERSE(SUBSTRING(REVERSE(subsequence), CHARINDEX('.', REVERSE(subsequence))+1, 100)))
		
		INSERT INTO #BOMCost
		SELECT m.item
			 , m.[Level]
			 , m.Parent
			 , m.oper_num
			 , m.sequence
			 , CAST(m.subsequence AS NVARCHAR(50))
			 , m.matl
			 , m.matl_qty
			 , m.matl_unit_cost + ISNULL(c.child_matl, 0)
			 , m.pi_process_cost + ISNULL(c.child_pi_process, 0)
			 , m.pi_resin_cost + ISNULL(c.child_pi_resin, 0)
			 , m.pi_hidden_profit + ISNULL(c.child_pi_hidden, 0)
			 , m.sf_labr_cost + ISNULL(c.child_sf_lbr, 0)
			 , m.sf_ovhd_cost + ISNULL(c.child_sf_ovhd, 0)
			 , m.fg_labr_cost + ISNULL(c.child_fg_lbr, 0)
			 , m.fg_ovhd_cost + ISNULL(c.child_fg_ovhd, 0)
		FROM #itemMatl as m
			LEFT OUTER JOIN childCost AS c
				ON 1 = (CASE WHEN @LevelCtr <> 0
								AND (CAST(m.[Level] AS NVARCHAR(3))+ '.' + CAST(m.oper_num AS NVARCHAR(3)) + '.' + CAST(m.sequence AS NVARCHAR(3))) = c.Parent
								AND c.subs_parent = m.subsequence
									THEN 1
							WHEN @LevelCtr = 0 AND c.subs_parent = m.subsequence
								THEN 1
							ELSE 0
							END)
		WHERE [Level] = (@LevelCtr - 1);		
		
		SELECT @LevelCtr = @LevelCtr - 1
		
	END
		UPDATE B
		SET matl_unit_cost = matl_unit_cost + (SELECT SUM(B2.matl_qty * B2.matl_unit_cost) FROM #BOMCost B2 WHERE B2.[Level] = 1)
		  , pi_process_cost = pi_process_cost + (SELECT SUM(B2.matl_qty * B2.pi_process_cost) FROM #BOMCost B2 WHERE B2.[Level] = 1)
		  , pi_resin_cost = pi_resin_cost + (SELECT SUM(B2.matl_qty * B2.pi_resin_cost) FROM #BOMCost B2 WHERE B2.[Level] = 1)
		  , pi_hidden_profit = pi_hidden_profit + (SELECT SUM(B2.matl_qty * B2.pi_hidden_profit) FROM #BOMCost B2 WHERE B2.[Level] = 1)
		  , sf_labr_cost = sf_labr_cost + (SELECT SUM(B2.matl_qty * B2.sf_labr_cost) FROM #BOMCost B2 WHERE B2.[Level] = 1)
		  , sf_ovhd_cost = sf_ovhd_cost + (SELECT SUM(B2.matl_qty * B2.sf_ovhd_cost) FROM #BOMCost B2 WHERE B2.[Level] = 1)
		  , fg_labr_cost = fg_labr_cost + (SELECT SUM(B2.matl_qty * B2.fg_labr_cost) FROM #BOMCost B2 WHERE B2.[Level] = 1)
		  , fg_ovhd_cost = fg_ovhd_cost + (SELECT SUM(B2.matl_qty * B2.fg_ovhd_cost) FROM #BOMCost B2 WHERE B2.[Level] = 1)
		FROM #BOMCost AS B
		WHERE b.[Level] = 0 AND Parent = 0
		
		SELECT *
		FROM #BOMCost
		ORDER BY subsequence, sequence, [Level]

	--SELECT @Item
	--	, 0 AS [Level]
	--	, SUM(matl_qty * matl_unit_cost) AS matl_cost
	--	, SUM(matl_qty * pi_process_cost) AS pi_process
	--	, SUM(matl_qty * pi_resin_cost) AS pi_resin
	--	, SUM(matl_qty * pi_hidden_profit) AS pi_hidden
	--	, SUM(matl_qty * sf_labr_cost)  AS sf_lbr
	--	, SUM(matl_qty * sf_ovhd_cost)  AS sf_ovhd
	--	, SUM(matl_qty * fg_labr_cost)  AS fg_lbr
	--	, SUM(matl_qty * fg_ovhd_cost)  AS fg_ovhd

	--FROM #BOMCost
	--WHERE [Level] = 1	
	

--	SELECT * FROM #itemMatl	
END