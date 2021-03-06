--LSP_DM_ActlCost_GetJobMatlTransCostingSp '19CL-00191',0,'SF-L4336','2020-01-04 06:03:13.000',60

--ALTER PROCEDURE LSP_DM_ActlCost_GetJobMatlTransCostingSp (
DECLARE  
	@Job					JobType		= '19-0002507'
  , @Suffix					SuffixType	= 0
  , @Item					ItemType	= 'FG-DK-100D'
  , @JobTransDate			DateType	= '2020-05-20'
  , @QtyTrans				QtyUnitType	= 200
 	--@Job					JobType		= '20-0000864'
  --, @Suffix					SuffixType	= 0
  --, @Item					ItemType	= 'FG-3RS2024'
  --, @JobTransDate			DateType	= '05/29/2020'
  --, @QtyTrans				QtyUnitType	= 10
 	--@Job					JobType		= '19-0002265'
  --, @Suffix					SuffixType	= 0
  --, @Item					ItemType	= 'FG-E21-211'
--) AS
BEGIN

	IF OBJECT_ID('tempdb..#itemMatl') IS NOT NULL
		DROP TABLE #itemMatl
	IF OBJECT_ID('tempdb..#subMatl') IS NOT NULL
		DROP TABLE #subMatl
	IF OBJECT_ID('tempdb..#SFMatl') IS NOT NULL
		DROP TABLE #SFMatl
	IF OBJECT_ID('tempdb..#ActualCost') IS NOT NULL
		DROP TABLE #ActualCost
		
	CREATE TABLE #itemMatl (
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
		
	DECLARE
		@Matl				ItemType
	  , @PrevMatl			ItemType
	  , @LotNo				LotType
	  , @TransDate			DateType
	  , @Level				INT
	  , @OperNum			OperNumType
	  , @Sequence			SequenceType
	  , @SubSequence		NVARCHAR(50)
	  , @CurrLevel			INT				= 1
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
		  , lot_no
		  , trans_date
		  , job_qty
		)	
	SELECT j.item
		 , @CurrLevel AS [Level]
		 , CAST(0 AS NVARCHAR(20)) AS Parent
		 , m.ref_release
		 , row_number() OVER (PARTITION BY j.item, m.ref_release ORDER BY m.ref_release ASC)
		 , CAST(m.ref_release AS NVARCHAR(5)) + '_' + CAST((row_number() OVER (PARTITION BY j.item, m.ref_release ORDER BY m.ref_release ASC)) AS NVARCHAR(50)) AS subsequence
		 , m.item
		 , SUM(m.qty * -1)
		 , m.lot
		 , MAX(m.trans_date)
		 , NULL
	FROM job AS j
		JOIN matltran AS m
			ON j.job = m.ref_num
			  AND j.suffix = m.ref_line_suf
			  AND m.ref_type = 'J'
			  AND m.trans_type IN ('I', 'W')
	WHERE j.job = @job
	  AND j.suffix = @Suffix
	  AND j.item = @Item	
	GROUP BY j.item, m.ref_release, m.item, m.lot
	HAVING SUM(m.qty * -1) <> 0.00
	
	UNION ALL
	SELECT @Item
		 , 0
		 , '0'
		 , '0'
		 , 0
		 , '0'
		 , item
		 , @QtyTrans
		 , @Job
		 , @JobTransDate
		 , j2.qty_released
	FROM job AS j2
	WHERE j2.job = @Job 
	  AND j2.suffix = @Suffix 
	  AND j2.item = @item
	ORDER BY m.ref_release
	
	--SELECT j.item
	--	 , @CurrLevel AS [Level]
	--	 , CAST(0 AS NVARCHAR(20)) AS Parent
	--	 , m.ref_release
	--	 , row_number() OVER (PARTITION BY j.item, m.ref_release ORDER BY m.ref_release ASC)
	--	 , CAST(m.ref_release AS NVARCHAR(5)) + '_' + CAST((row_number() OVER (PARTITION BY j.item, m.ref_release ORDER BY m.ref_release ASC)) AS NVARCHAR(50)) AS subsequence
	--	 , m.item
	--	 , SUM(m.qty * -1)
	--	 , m.lot
	--	 , MAX(m.trans_date)
	--	 , NULL
	--FROM job AS j
	--	JOIN matltran AS m
	--		ON j.job = m.ref_num
	--		  AND j.suffix = m.ref_line_suf
	--		  AND m.ref_type = 'J'
	--		  AND m.trans_type IN ('I', 'W')
	--WHERE j.job = @job
	--  AND j.suffix = @Suffix
	--  AND j.item = @Item
	--GROUP BY j.item, m.ref_release, m.item, m.lot	
	--SELECT * FROM job WHERE job = @Job AND suffix = @Suffix AND item = @item
	
	SELECT item
		 , [Level]
		 , Parent
		 , oper_num
		 , sequence
		 , subsequence
		 , matl
		 , lot_no
		 , trans_date
		 , matl_qty
		 
	INTO #SFMatl
	FROM #itemMatl
	WHERE (matl LIKE 'SF-%' OR matl LIKE 'FG-%')
	  AND [Level] > 0
	ORDER BY item
	
	SELECT item
		 , [Level]
		 , CAST(Parent AS NVARCHAR(50)) AS Parent
		 , oper_num
		 , sequence
		 , CAST(subsequence AS NVARCHAR(50)) AS subsequence
		 , matl
		 , CAST(lot_no AS NVARCHAR(50)) AS lot_no
		 , trans_date
		 , matl_qty
	INTO #subMatl
	FROM #itemMatl
	WHERE (matl LIKE 'SF-%' OR matl LIKE 'FG-%')
	  AND [Level] > 0	  
	ORDER BY item
	
	WHILE (SELECT COUNT(*) FROM #SFMatl) > 0
	BEGIN
		--SELECT * FROM #SFMatl
		--SELECT * FROM #subMatl
		
		SELECT @CurrLevel = @CurrLevel + 1
		TRUNCATE TABLE #subMatl

		DECLARE matlCrsr CURSOR FAST_FORWARD FOR	
		SELECT matl
			 , [Level]
			 , oper_num
			 , sequence
			 , subsequence
			 , lot_no
			 , trans_date
		FROM #SFMatl
		WHERE matl LIKE 'SF-%' OR matl LIKE 'FG-%'
		ORDER BY [Level], matl
		
		OPEN matlCrsr
		FETCH FROM matlCrsr INTO
			@Matl
		  , @Level
		  , @OperNum
		  , @Sequence
		  , @SubSequence
		  , @LotNo
		  , @TransDate
		
		WHILE (@@FETCH_STATUS = 0)
		BEGIN					
			
			IF EXISTS (SELECT * FROM job WHERE job = @LotNo)
			BEGIN			
				
				--SELECT j.item
				--	 , @CurrLevel AS [Level]
				--	 , CAST(@Level AS NVARCHAR(5)) + '.' + CAST(@OperNum AS NVARCHAR(5)) + '.' + CAST(@Sequence AS NVARCHAR(5))AS Parent
				--	 , m.ref_release
				--	 , row_number() OVER (PARTITION BY j.item, m.ref_release ORDER BY m.ref_release ASC)
				--	, @SubSequence + '.' + CAST((row_number() OVER (PARTITION BY j.item, m.ref_release ORDER BY m.ref_release ASC)) AS NVARCHAR(50)) AS subsequence
				--	 , m.item
				--	 , SUM(m.qty * -1)
				--	 , m.lot
				--	 , MAX(m.trans_date)
				--FROM job AS j
				--	JOIN matltran AS m
				--		ON j.job = m.ref_num
				--		  AND j.suffix = m.ref_line_suf
				--		  AND m.ref_type = 'J'
				--		  AND m.trans_type IN ('I', 'W')
				--WHERE j.job = @LotNo
				--  AND j.suffix = 0
				--  AND j.item = @Matl
				--GROUP BY j.item, m.ref_release, m.item, m.lot
				--ORDER BY m.ref_release, row_number() OVER (PARTITION BY j.item, m.ref_release ORDER BY m.ref_release ASC)

				INSERT INTO #itemMatl 
					(	item
					  , [Level]
					  , Parent
					  , oper_num
					  , sequence
					  , subsequence
					  , matl
					  , matl_qty
					  , lot_no
					  , trans_date
					)	
				
				SELECT j.item
					 , @CurrLevel AS [Level]
					 , CAST(@Level AS NVARCHAR(5)) + '.' + CAST(@OperNum AS NVARCHAR(5)) + '.' + CAST(@Sequence AS NVARCHAR(5))AS Parent
					 , m.ref_release
					 , row_number() OVER (PARTITION BY j.item, m.ref_release ORDER BY m.ref_release ASC)
					, @SubSequence + '.' + CAST((row_number() OVER (PARTITION BY j.item, m.ref_release ORDER BY m.ref_release ASC)) AS NVARCHAR(50)) AS subsequence
					 , m.item
					 , SUM(m.qty * -1)
					 , m.lot
					 , MAX(m.trans_date)
				FROM job AS j
					JOIN matltran AS m
						ON j.job = m.ref_num
						  AND j.suffix = m.ref_line_suf
						  AND m.ref_type = 'J'
						  AND m.trans_type IN ('I', 'W')
				WHERE j.job = @LotNo
				  AND j.suffix = 0
				  AND j.item = @Matl
				GROUP BY j.item, m.ref_release, m.item, m.lot
				ORDER BY m.ref_release, row_number() OVER (PARTITION BY j.item, m.ref_release ORDER BY m.ref_release ASC);

--SELECT @CurrLevel, @LotNo, @Matl, @Level, @Sequence
--SELECT * FROM #itemMatl
--ORDER BY subsequence, sequence, [Level]
--SELECT @Matl, @Level, @OperNum, @Sequence, @SubSequence, @LotNo, @TransDate, @CurrLevel;

				
				WITH CTE_sub AS (			
				SELECT j.item
					 , @Level + 1 AS [Level]
					 , CAST(@Level AS NVARCHAR(5)) + '.' + CAST(@OperNum AS NVARCHAR(5)) + '.' + CAST(@Sequence AS NVARCHAR(5))AS Parent
					 , m.ref_release AS oper_num
					 , row_number() OVER (PARTITION BY j.item, m.ref_release ORDER BY m.ref_release ASC) AS sequence
					, @SubSequence + '.' + CAST((row_number() OVER (PARTITION BY j.item, m.ref_release ORDER BY m.ref_release ASC)) AS NVARCHAR(50)) AS subsequence
					 , m.item AS matl
					 , m.lot
					 , MAX(m.trans_date) AS trans_date
					 , SUM(m.qty * -1) AS qty
					 
				FROM job AS j
					JOIN matltran AS m
						ON j.job = m.ref_num
						  AND j.suffix = m.ref_line_suf
						  AND m.ref_type = 'J'
						  AND m.trans_type IN ('I', 'W')
				WHERE j.job = @LotNo
				  AND j.suffix = 0
				  --AND (m.item LIKE 'SF-%' OR m.item LIKE 'FG-%')
				 -- AND m.lot <> @LotNo
				GROUP BY j.item, m.ref_release, m.item, m.lot)
				--ORDER BY m.ref_release, row_number() OVER (PARTITION BY j.item, m.ref_release ORDER BY m.ref_release ASC) 
				
				INSERT INTO #submatl
				SELECT *
				FROM CTE_sub
				WHERE (item LIKE 'SF-%' OR item LIKE 'FG-%')
				ORDER BY oper_num ;
				
			END
		
			FETCH NEXT FROM matlCrsr INTO
				@Matl
			  , @Level
			  , @OperNum
			  , @Sequence
			  , @SubSequence
			  , @LotNo
			  , @TransDate
		
		END
		
		CLOSE matlCrsr
		DEALLOCATE matlCrsr
		
		TRUNCATE TABLE #SFMatl
		
		INSERT INTO #SFMatl
		SELECT * FROM #subMatl	
		
		--SELECT * FROM #subMatl	
	END
	
	--SELECT * FROM #itemMatl
	--ORDER BY subsequence, sequence, [Level]
	
	--/*******
	DECLARE matlCrsr CURSOR FAST_FORWARD FOR
	SELECT DISTINCT matl
		 , lot_no
		 , trans_date
	FROM #itemMatl
	
	OPEN matlCrsr
	FETCH FROM matlCrsr INTO
		@Matl
	  , @LotNo
	  , @TransDate
		
	WHILE (@@FETCH_STATUS = 0)
	BEGIN
		SELECT @matl_unit_cost_usd		= 0
			 , @matl_landed_cost_usd	= 0
			 , @pi_fg_process_usd		= 0
			 , @pi_resin_usd			= 0
			 , @pi_vend_cost_usd		= 0
			 , @pi_hidden_profit_usd	= 0
			 , @sf_lbr_cost_usd			= 0
			 , @sf_ovhd_cost_usd		= 0
			 , @fg_lbr_cost_usd			= 0
			 , @fg_ovhd_cost_usd		= 0
			 , @matl_unit_cost_php		= 0
			 , @matl_landed_cost_php	= 0
			 , @pi_fg_process_php		= 0
			 , @pi_resin_php			= 0
			 , @pi_vend_cost_php		= 0
			 , @pi_hidden_profit_php	= 0
			 , @sf_lbr_cost_php			= 0
			 , @sf_ovhd_cost_php		= 0
			 , @fg_lbr_cost_php			= 0
			 , @fg_ovhd_cost_php		= 0
		
		EXEC dbo.LSP_ActlCost_GetMatlCostingSp 
					@Matl, @LotNo, @TransDate
				  , @JobQty OUTPUT
				  , @matl_unit_cost_usd OUTPUT, @matl_landed_cost_usd OUTPUT
				  , @pi_fg_process_usd OUTPUT, @pi_resin_usd OUTPUT, @pi_vend_cost_usd OUTPUT, @pi_hidden_profit_usd OUTPUT
				  , @sf_lbr_cost_usd OUTPUT, @sf_ovhd_cost_usd OUTPUT
				  , @fg_lbr_cost_usd OUTPUT, @fg_ovhd_cost_usd OUTPUT
				  , @matl_unit_cost_php OUTPUT, @matl_landed_cost_php OUTPUT
				  , @pi_fg_process_php OUTPUT, @pi_resin_php OUTPUT, @pi_vend_cost_php OUTPUT, @pi_hidden_profit_php OUTPUT
				  , @sf_lbr_cost_php OUTPUT, @sf_ovhd_cost_php OUTPUT
				  , @fg_lbr_cost_php OUTPUT, @fg_ovhd_cost_php OUTPUT
		
		UPDATE #itemMatl
		SET job_qty = ISNULL(@JobQty, 0)
		  , matl_unit_cost_usd = ISNULL(@matl_unit_cost_usd, 0) * matl_qty
		  , matl_landed_cost_usd = ISNULL(@matl_landed_cost_usd, 0) * matl_qty
		  , pi_fg_process_usd = ISNULL(@pi_fg_process_usd, 0) * matl_qty
		  , pi_resin_usd = ISNULL(@pi_resin_usd, 0) * matl_qty
		  , pi_vend_cost_usd = ISNULL(@pi_vend_cost_usd, 0) * matl_qty
		  , pi_hidden_profit_usd = ISNULL(@pi_hidden_profit_usd, 0) * matl_qty
		  , sf_lbr_cost_usd = ISNULL(@sf_lbr_cost_usd, 0) * matl_qty
		  , sf_ovhd_cost_usd = ISNULL(@sf_ovhd_cost_usd, 0) * matl_qty
		  , fg_lbr_cost_usd = ISNULL(@fg_lbr_cost_usd, 0) * matl_qty
		  , fg_ovhd_cost_usd = ISNULL(@fg_ovhd_cost_usd, 0) * matl_qty
		  , matl_unit_cost_php = ISNULL(@matl_unit_cost_php, 0) * matl_qty
		  , matl_landed_cost_php = ISNULL(@matl_landed_cost_php, 0) * matl_qty
		  , pi_fg_process_php = ISNULL(@pi_fg_process_php, 0) * matl_qty
		  , pi_resin_php = ISNULL(@pi_resin_php, 0) * matl_qty
		  , pi_vend_cost_php = ISNULL(@pi_vend_cost_php, 0) * matl_qty
		  , pi_hidden_profit_php = ISNULL(@pi_hidden_profit_php, 0) * matl_qty
		  , sf_lbr_cost_php = ISNULL(@sf_lbr_cost_php, 0) * matl_qty
		  , sf_ovhd_cost_php = ISNULL(@sf_ovhd_cost_php, 0) * matl_qty
		  , fg_lbr_cost_php = ISNULL(@fg_lbr_cost_php, 0) * matl_qty
		  , fg_ovhd_cost_php = ISNULL(@fg_ovhd_cost_php, 0) * matl_qty
		WHERE matl = @Matl
		  AND lot_no = @LotNo
		  AND trans_date = @TransDate
		  
		FETCH NEXT FROM matlCrsr INTO
			@Matl
		  , @LotNo
		  , @TransDate
	
	END
	
	CLOSE matlCrsr
	DEALLOCATE matlCrsr
	--*****/
	SELECT @LevelCtr = MAX([Level])
	FROM #itemMatl
	
	SELECT * 
	INTO #ActualCost
	FROM #itemMatl
	WHERE [Level] = @LevelCtr
		
	WHILE @LevelCtr >= 0
	BEGIN	
		-- SELECT @LevelCtr
		--SELECT Parent, REVERSE(SUBSTRING(REVERSE(subsequence), CHARINDEX('.', REVERSE(subsequence))+1, 100))
		--, CAST( REVERSE(SUBSTRING(REVERSE(subsequence), CHARINDEX('.', REVERSE(subsequence))+1, 100)) AS NVARcHAR(50))AS subs_parent
		--, *
		--FROM #ActualCost
		--WHERE [Level] = @LevelCtr;
	
		WITH childCost AS (
		SELECT Parent, CAST( REVERSE(SUBSTRING(REVERSE(subsequence), CHARINDEX('.', REVERSE(subsequence))+1, 100)) AS NVARcHAR(50))AS subs_parent
			--, lot_no
			--, trans_date
			, SUM(matl_unit_cost_usd ) AS child_matl_usd
			, SUM(matl_landed_cost_usd) AS child_matl_landed_usd
			, SUM(pi_fg_process_usd) AS child_pi_fg_process_usd
			, SUM(pi_resin_usd) AS child_pi_resin_usd
			, SUM(pi_vend_cost_usd) AS child_pi_vend_cost_usd
			, SUM(pi_hidden_profit_usd) AS child_pi_hidden_usd
			, SUM(sf_lbr_cost_usd)  AS child_sf_lbr_usd
			, SUM(sf_ovhd_cost_usd)  AS child_sf_ovhd_usd
			, SUM(fg_lbr_cost_usd)  AS child_fg_lbr_usd
			, SUM(fg_ovhd_cost_usd)  AS child_fg_ovhd_usd

			, SUM(matl_unit_cost_php) AS child_matl_php
			, SUM(matl_landed_cost_php) AS child_matl_landed_php
			, SUM(pi_fg_process_php) AS child_pi_fg_process_php
			, SUM(pi_resin_php) AS child_pi_resin_php
			, SUM(pi_vend_cost_php) AS child_pi_vend_cost_php
			, SUM(pi_hidden_profit_php) AS child_pi_hidden_php
			, SUM(sf_lbr_cost_php)  AS child_sf_lbr_php
			, SUM(sf_ovhd_cost_php)  AS child_sf_ovhd_php
			, SUM(fg_lbr_cost_php)  AS child_fg_lbr_php
			, SUM(fg_ovhd_cost_php)  AS child_fg_ovhd_php
		FROM #ActualCost
		WHERE [Level] = @LevelCtr		
		GROUP BY Parent, REVERSE(SUBSTRING(REVERSE(subsequence), CHARINDEX('.', REVERSE(subsequence))+1, 100))
		)
		
		INSERT INTO #ActualCost
		SELECT m.item
			 , m.[Level]
			 , m.Parent
			 , m.oper_num
			 , m.sequence
			 , CAST(m.subsequence AS NVARCHAR(50))
			 , m.matl
			 , m.matl_qty
			 , m.lot_no
			 , m.trans_date
			 , m.job_qty
		     , (m.matl_unit_cost_usd) + ISNULL(c.child_matl_usd / m.job_qty, 0) * matl_qty
			 , (m.matl_landed_cost_usd) + ISNULL(c.child_matl_landed_usd / m.job_qty, 0) * matl_qty
			 , (m.pi_fg_process_usd) + ISNULL(c.child_pi_fg_process_usd / m.job_qty, 0) * matl_qty
			 , (m.pi_resin_usd) + ISNULL(c.child_pi_resin_usd / m.job_qty, 0) * matl_qty
			 , (m.pi_vend_cost_usd) + ISNULL(c.child_pi_vend_cost_usd / m.job_qty, 0) * matl_qty
			 , (m.pi_hidden_profit_usd) + ISNULL(c.child_pi_hidden_usd / m.job_qty, 0) * matl_qty
			 , (m.sf_lbr_cost_usd) + ISNULL(c.child_sf_lbr_usd / m.job_qty, 0) * matl_qty
			 , (m.sf_ovhd_cost_usd) + ISNULL(c.child_sf_ovhd_usd / m.job_qty, 0) * matl_qty
			 , (m.fg_lbr_cost_usd) + ISNULL(c.child_fg_lbr_usd / m.job_qty, 0) * matl_qty
			 , (m.fg_ovhd_cost_usd) + ISNULL(c.child_fg_ovhd_usd / m.job_qty, 0) * matl_qty
			 , (m.matl_unit_cost_php) + ISNULL(c.child_matl_php / m.job_qty, 0) * matl_qty
			 , (m.matl_landed_cost_php) + ISNULL(c.child_matl_landed_php / m.job_qty, 0) * matl_qty
			 , (m.pi_fg_process_php) + ISNULL(c.child_pi_fg_process_php / m.job_qty, 0) * matl_qty
			 , (m.pi_resin_php) + ISNULL(c.child_pi_resin_php / m.job_qty, 0) * matl_qty
			 , (m.pi_vend_cost_php) + ISNULL(c.child_pi_vend_cost_php / m.job_qty, 0) * matl_qty
			 , (m.pi_hidden_profit_php) + ISNULL(c.child_pi_hidden_php / m.job_qty, 0) * matl_qty
			 , (m.sf_lbr_cost_php) + ISNULL(c.child_sf_lbr_php / m.job_qty, 0) * matl_qty
			 , (m.sf_ovhd_cost_php) + ISNULL(c.child_sf_ovhd_php / m.job_qty, 0) * matl_qty
			 , (m.fg_lbr_cost_php) + ISNULL(c.child_fg_lbr_php / m.job_qty, 0) * matl_qty
			 , (m.fg_ovhd_cost_php) + ISNULL(c.child_fg_ovhd_php / m.job_qty, 0) * matl_qty
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
		SET matl_unit_cost_usd = ISNULL(matl_unit_cost_usd,0) + (SELECT SUM(B2.matl_unit_cost_usd) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , matl_landed_cost_usd = ISNULL(matl_landed_cost_usd,0) + (SELECT SUM(B2.matl_landed_cost_usd) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , pi_fg_process_usd = ISNULL(pi_fg_process_usd,0) + (SELECT SUM(B2.pi_fg_process_usd) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , pi_resin_usd = ISNULL(pi_resin_usd,0) + (SELECT SUM(B2.pi_resin_usd) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , pi_vend_cost_usd = ISNULL(pi_vend_cost_usd,0) + (SELECT SUM(B2.pi_vend_cost_usd) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , pi_hidden_profit_usd = ISNULL(pi_hidden_profit_usd,0) + (SELECT SUM(B2.pi_hidden_profit_usd) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , sf_lbr_cost_usd = ISNULL(sf_lbr_cost_usd,0) + (SELECT SUM(B2.sf_lbr_cost_usd) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , sf_ovhd_cost_usd = ISNULL(sf_ovhd_cost_usd,0) + (SELECT SUM(B2.sf_ovhd_cost_usd) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , fg_lbr_cost_usd = ISNULL(fg_lbr_cost_usd,0) + (SELECT SUM(B2.fg_lbr_cost_usd) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , fg_ovhd_cost_usd = ISNULL(fg_ovhd_cost_usd,0) + (SELECT SUM(B2.fg_ovhd_cost_usd) FROM #ActualCost B2 WHERE B2.[Level] = 1)

		  , matl_unit_cost_php = ISNULL(matl_unit_cost_php,0) + (SELECT SUM(B2.matl_unit_cost_php) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , matl_landed_cost_php = ISNULL(matl_landed_cost_php,0) + (SELECT SUM(B2.matl_landed_cost_php) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , pi_fg_process_php = ISNULL(pi_fg_process_php,0) + (SELECT SUM(B2.pi_fg_process_php) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , pi_resin_php = ISNULL(pi_resin_php,0) + (SELECT SUM(B2.pi_resin_php) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , pi_vend_cost_php = ISNULL(pi_vend_cost_php,0) + (SELECT SUM(B2.pi_vend_cost_php) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , pi_hidden_profit_php = ISNULL(pi_hidden_profit_php,0) + (SELECT SUM(B2.pi_hidden_profit_php) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , sf_lbr_cost_php = ISNULL(sf_lbr_cost_php,0) + (SELECT SUM(B2.sf_lbr_cost_php) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , sf_ovhd_cost_php = ISNULL(sf_ovhd_cost_php,0) + (SELECT SUM(B2.sf_ovhd_cost_php) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , fg_lbr_cost_php = ISNULL(fg_lbr_cost_php,0) + (SELECT SUM(B2.fg_lbr_cost_php) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		  , fg_ovhd_cost_php = ISNULL(fg_ovhd_cost_php,0) + (SELECT SUM(B2.fg_ovhd_cost_php) FROM #ActualCost B2 WHERE B2.[Level] = 1)
		FROM #ActualCost AS B
		WHERE b.[Level] = 0 AND Parent = 0

	--/*******
	SELECT --CAST(m.matl_unit_cost_usd / m.matl_qty AS DECIMAL(18,8)) AS matl_unit_cost_usd
		 --, CAST(m.matl_landed_cost_usd / m.matl_qty  AS DECIMAL(18,8)) AS matl_landed_cost_usd, 
		 A.* 
	FROM #ActualCost AS A
		--LEFT OUTER JOIN #itemMatl AS m ON a.item = m.item AND A.[Level] = m.[Level] AND A.matl = m.matl AND A.subsequence = m.subsequence
	--WHERE a.Level = 1 --AND a.matl = 'SF-MMU003'
	ORDER BY subsequence, CAST(sequence AS INT), [Level]
	--****/

END
