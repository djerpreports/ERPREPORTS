ALTER PROCEDURE LSP_Rpt_NewDM_WIPShopFloorReportSp

AS
BEGIN

	IF OBJECT_ID('tempdb..#WIPShopFloor') IS NOT NULL
		DROP TABLE #WIPShopFloor
	
	DECLARE
		@Job					JobType
	  , @Suffix					SuffixType
	  , @Item					ItemType
	  , @Stat					JobStatusType
	  , @QtyReleased			QtyUnitType
	  , @QtyCompleted			QtyUnitType
	  , @QtyScrapped			QtyUnitType
	  , @JobStartDate			DateType
	  , @QtyWip					QtyUnitType
	 
	CREATE TABLE #WIPShopFloor(
		TransDate				DATETIME
	  , Item					NVARCHAR(100)
	  , ItemDesc				NVARCHAR(100)
	  , JOReference				NVARCHAR(50)
	  , WIPQty					DECIMAL(18,8)
	  , ItemLot					NVARCHAR(50)
	  , MatlUnit_PHP			DECIMAL(18,8)
	  , LandedUnit_PHP			DECIMAL(18,8)
	  , PIFGProcessUnit_PHP		DECIMAL(18,8)
	  , PIResinUnit_PHP			DECIMAL(18,8)
	  , PIHiddenUnit_PHP		DECIMAL(18,8)
	  , SFAddedUnit_PHP			DECIMAL(18,8)
	  , FGAddedUnit_PHP			DECIMAL(18,8)
	  , MatlUnit_USD			DECIMAL(18,8)
	  , LandedUnit_USD			DECIMAL(18,8)
	  , PIFGProcessUnit_USD		DECIMAL(18,8)
	  , PIResinUnit_USD			DECIMAL(18,8)
	  , PIHiddenUnit_USD		DECIMAL(18,8)
	  , SFAddedUnit_USD			DECIMAL(18,8)
	  , FGAddedUnit_USD			DECIMAL(18,8)
	)

	DECLARE wipJobCrsr CURSOR FAST_FORWARD FOR
	SELECT j.job
		 , j.suffix
		 , j.item
		 , j.stat
		 , j.qty_released
		 , j.qty_complete
		 , j.qty_scrapped
		 , js.[start_date]
	FROM job AS j
		LEFT OUTER JOIN job_sch AS js
			ON j.job = js.job
			  AND j.suffix = js.suffix
	WHERE j.stat = 'R'  
	 AND (j.qty_complete + j.qty_scrapped) <> j.qty_released  
--	 AND j.qty_complete > 0

	OPEN wipJobCrsr
	FETCH FROM wipJobCrsr INTO
		@Job
	  , @Suffix
	  , @Item
	  , @Stat
	  , @QtyReleased
	  , @QtyCompleted
	  , @QtyScrapped
	  , @JobStartDate
	
	WHILE @@FETCH_STATUS = 0
	BEGIN
	
		SET @QtyWip = (@QtyReleased - (@QtyCompleted + @QtyScrapped) )
	
		EXEC dbo.LSP_DM_ActlCost_GetJobMatlTransCostingSp @job, @Suffix, @Item, @JobStartDate, @QtyWip
		
		PRINT (@job+ @Item)
		
		INSERT INTO #WIPShopFloor
		SELECT ac.trans_date
			 , ac.matl
			 , i.[description]
			 , @Job + ' - ' + RIGHT( ('00' +CAST(@Suffix AS NVARCHAR(5))),2 )
			 , ac.matl_qty - ((ac.matl_qty / @QtyReleased) * (@QtyCompleted + @QtyScrapped))--@QtyWip
			 , ac.lot_no
			 
			 , ac.matl_unit_cost_php / ac.matl_qty AS MatlUnit_PHP
			 , ac.matl_landed_cost_php / ac.matl_qty 
			 , ac.pi_fg_process_php / ac.matl_qty 
			 , ac.pi_resin_php / ac.matl_qty 
			 , ac.pi_hidden_profit_php / ac.matl_qty 
			 , (ac.sf_lbr_cost_php / ac.matl_qty) + (ac.sf_ovhd_cost_php / ac.matl_qty)
			 , (ac.fg_lbr_cost_php / ac.matl_qty) + (ac.fg_ovhd_cost_php / ac.matl_qty)
			 
			 , ac.matl_unit_cost_usd / ac.matl_qty AS MatlUnit_usd
			 , ac.matl_landed_cost_usd / ac.matl_qty 
			 , ac.pi_fg_process_usd / ac.matl_qty 
			 , ac.pi_resin_usd / ac.matl_qty 
			 , ac.pi_hidden_profit_usd / ac.matl_qty 
			 , (ac.sf_lbr_cost_usd / ac.matl_qty) + (ac.sf_ovhd_cost_usd / ac.matl_qty)
			 , (ac.fg_lbr_cost_usd / ac.matl_qty) + (ac.fg_ovhd_cost_usd / ac.matl_qty)
			 
		FROM ##ActualCost AS ac
			LEFT OUTER JOIN item AS i
				ON ac.matl = i.item
		WHERE [Level] = 1
		ORDER BY subsequence, sequence, [Level]

		FETCH NEXT FROM wipJobCrsr INTO
			@Job
		  , @Suffix
		  , @Item
		  , @Stat
		  , @QtyReleased
		  , @QtyCompleted
		  , @QtyScrapped
		  , @JobStartDate
	
	END
	
	CLOSE wipJobCrsr
	DEALLOCATE wipJobCrsr	


	SELECT *
		 , (MatlUnit_PHP + LandedUnit_PHP 
				+ PIFGProcessUnit_PHP + PIResinUnit_PHP + PIHiddenUnit_PHP 
				+ SFAddedUnit_PHP + FGAddedUnit_PHP) 
			* WIPQty AS TotalWIPCost_PHP
		 , (MatlUnit_USD + LandedUnit_USD 
				+ PIFGProcessUnit_USD + PIResinUnit_USD + PIHiddenUnit_USD 
				+ SFAddedUnit_USD + FGAddedUnit_USD) 
			* WIPQty AS TotalWIPCost_USD
	FROM #WIPShopFloor
	
	
END