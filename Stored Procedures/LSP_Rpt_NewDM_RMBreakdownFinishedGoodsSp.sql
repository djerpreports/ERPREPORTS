--CREATE PROCEDURE LSP_Rpt_NewDM_RMBreakdownFinishedGoodsSp (
DECLARE
	@StartDate				DateType = '05/01/2020'
  , @EndDate				DateType = '05/31/2020'
--) AS
BEGIN

	IF OBJECT_ID('tempdb..#RMBreakDownFG') IS NOT NULL
		DROP TABLE #RMBreakDownFG

	DECLARE @ship_tran AS TABLE (
		TransDate			DateType
	  , Item				ItemType
	  , QtyShipped			QtyUnitType
	  , JobOrder			JobType
	  , JobSuffix			SuffixType
	  , PONumber			NVARCHAR(20)
	  
	)

	DECLARE @FGReceipts AS TABLE (
		TransDate			DateType
	  , Item				ItemType
	  , QtyCompleted		QtyUnitType
	  , JobOrder			JobType
	  , JobSuffix			SuffixType
	  , PONumber			NVARCHAR(20)	  
	)

	CREATE TABLE #RMBreakDownFG (
		JONum				NVARCHAR(20)
	  , PONum				NVARCHAR(20)
	  , Item				NVARCHAR(60)
	  , matl				nvarchar(60)
	  , matl_desc			NVARCHAR(100)
	  , StdLbrHrs			decimal(18, 10) 
	  , ActlLbrHrs			decimal(18, 10) 
	  , std_matl_unit		DECIMAL(18, 8)
	  , std_process_unit	DECIMAL(18, 8)
	  , pi_resin_unit		DECIMAL(18, 8)
	  , pi_hidden_unit		DECIMAL(18, 8)
	  , sf_lbr_unit			DECIMAL(18, 8)
	  , sf_ovhd_unit		DECIMAL(18, 8)
	  , fg_lbr_unit			DECIMAL(18, 8)
	  , fg_ovhd_unit		DECIMAL(18, 8)
	  , total_std_unit		decimal(18, 8) 
	  , [Level]				int 
	  , sequence			nvarchar(3) 
	  , subsequence			nvarchar(50) 
	  , lot_no				nvarchar(50) 
	  , matl_qty			decimal(18, 8) 
	  , job_qty				bigint 
	  , job_matl_qty		decimal(18, 8) 
	  , actl_matl_qty		decimal(18, 8) 
	  , matl_unit_cost_php	decimal(18, 8) 
	  , matl_landed_cost_php	decimal(18, 8) 
	  , pi_fg_process_php	decimal(18, 8) 
	  , pi_resin_php		decimal(18, 8) 
	  , pi_hidden_profit_php	decimal(18, 8) 
	  , sf_lbr_cost_php		decimal(18, 8) 
	  , sf_ovhd_cost_php	decimal(18, 8) 
	  , fg_lbr_cost_php		decimal(18, 8) 
	  , fg_ovhd_cost_php	decimal(18, 8) 
	  , total_actl_unit		decimal(18, 8) 
	  , nolanded_actl_unit	decimal(18, 8) 
	)

	DECLARE
		@TransDate				DateType
	  , @Item					ItemType
	  , @QtyCompleted			INT
	  , @JobOrder				JobType
	  , @JobSuffix				SuffixType
	  , @PONum					NVARCHAR(20)
	  , @SQLStr					NVARCHAR(1000)
	  
	INSERT INTO @ship_tran
	SELECT m.trans_date
	  , m.item
	  , m.qty
	  , (SELECT TOP(1) matltran2.ref_num
		  FROM matltran matltran2
		  WHERE m.lot = matltran2.lot AND matltran2.trans_type = 'F' AND m.item = matltran2.item
		  ORDER BY matltran2.trans_date DESC)
	  , (SELECT TOP(1) matltran2.ref_line_suf
	  FROM matltran matltran2
	  WHERE m.lot = matltran2.lot AND matltran2.trans_type = 'F' AND m.item = matltran2.item
	  ORDER BY matltran2.trans_date DESC)
	  , coi.Uf_ponum	  
	FROM matltran AS m
		JOIN item AS i
			ON m.item = i.item
		JOIN coitem AS coi
			ON m.ref_num = coi.co_num AND m.ref_line_suf = coi.co_line

	WHERE m.trans_date BETWEEN dbo.MidnightOf(@StartDate) AND dbo.DayEndOf(@EndDate)
	  AND m.trans_type = 'S' AND m.ref_type = 'O'
	  AND m.item LIKE 'FG-%';
	
	WITH CTE_ship AS
	(SELECT MAX(TransDate) AS TransDate
	  , Item
	  , (SUM(QtyShipped) * (-1)) AS QtyShipped
	  , JobOrder
	  , JobSuffix
	  , PONumber
	  
	FROM @ship_tran
	GROUP BY PONumber, Item, JobOrder, JobSuffix)
	
	INSERT INTO @FGReceipts
	SELECT m.trans_date
	  , m.item
	  , m.qty
	  , m.ref_num
	  , m.ref_line_suf
	  , coi.Uf_ponum
	  
	FROM matltran AS m
		JOIN item AS i
			ON m.item = i.item
		JOIN job AS j
			ON m.ref_num = j.job AND m.ref_line_suf = j.suffix
		LEFT OUTER JOIN coitem AS coi
			ON j.ord_num = coi.co_num AND j.ord_line = coi.co_line
		LEFT OUTER JOIN CTE_ship AS ship
			ON coi.Uf_ponum = ship.PONumber AND m.ref_num = ship.JobOrder AND m.ref_line_suf = ship.JobSuffix
	
	WHERE m.trans_date BETWEEN dbo.MidnightOf(@StartDate) AND dbo.DayEndOf(@EndDate)
	  AND m.qty > 0 AND m.trans_type = 'F' AND m.ref_type = 'J'
	  AND m.item LIKE 'FG-%';
	  
	DECLARE FGCrsr CURSOR FAST_FORWARD FOR
	SELECT MAX(TransDate) AS TransDate
	  , PONumber
	  , JobOrder
	  , JobSuffix
	  , Item
	  , SUM(QtyCompleted)
	  
	FROM @FGReceipts
	--WHERE JobOrder IN ('20-0000864', '20-0000859','20-0000321')
	GROUP BY JobOrder, JobSuffix, PONumber, Item
	ORDER BY MAX(TransDate)

	--SELECT * 
	--FROM @FGReceipts

	OPEN FGCrsr
	FETCH FROM FGCrsr INTO
		@TransDate
	  , @PONum
	  , @JobOrder
	  , @JobSuffix
	  , @Item
	  , @QtyCompleted
	
	WHILE @@FETCH_STATUS = 0
	BEGIN
		
				
		SET @SQLStr = 'SELECT * 
						FROM OPENROWSET(''SQLNCLI'', ''Server=SYTELINEERP803;Database=LSPI803_App;UID=sa;Pwd=P@$$w0rd'',
						''SET FMTONLY OFF;SET NOCOUNT ON; EXEC dbo.LSP_Rpt_NewDM_RMBreakdownPerJOSp ''''' + @JobOrder + ''''', ''''' + @PONum + ''''', '+ CAST(@QtyCompleted AS NVARCHAR(10))  + ' '') AS a'
		--PRINT @JobOrder

		INSERT INTO #RMBreakDownFG ( 
			JONum
		  , PONum
		  , matl
		  , matl_desc
		  , StdLbrHrs
		  , ActlLbrHrs
		  , std_matl_unit
		  , std_process_unit
		  , pi_resin_unit
		  , pi_hidden_unit
		  , sf_lbr_unit
		  , sf_ovhd_unit
		  , fg_lbr_unit
		  , fg_ovhd_unit
		  , total_std_unit
		  , [Level]
		  , [sequence]
		  , subsequence
		  , lot_no
		  , matl_qty
		  , job_qty
		  , job_matl_qty
		  , actl_matl_qty
		  , matl_unit_cost_php
		  , matl_landed_cost_php
		  , pi_fg_process_php
		  , pi_resin_php
		  , pi_hidden_profit_php
		  , sf_lbr_cost_php
		  , sf_ovhd_cost_php
		  , fg_lbr_cost_php
		  , fg_ovhd_cost_php
		  , total_actl_unit
		  , nolanded_actl_unit
		)
		EXECUTE sp_executesql @SQLStr

		UPDATE #RMBreakDownFG
		SET Item = @Item
		WHERE JONum = @JobOrder
		  AND PONum = @PONum

		FETCH NEXT FROM FGCrsr INTO
			@TransDate
		  , @PONum
		  , @JobOrder
		  , @JobSuffix
		  , @Item
		  , @QtyCompleted
	END
	
	CLOSE FGCrsr
	DEALLOCATE FGCrsr
	
	SELECT JONum
		 , PONum
		 , Item
		 , matl
		 , matl_desc
		 , actl_matl_qty
		 , std_matl_unit
		 , std_process_unit
		 , pi_resin_unit
		 , pi_hidden_unit
		 , sf_lbr_unit
		 , sf_ovhd_unit
		 , fg_lbr_unit
		 , fg_ovhd_unit
		 , total_std_unit
		 , matl_unit_cost_php
		 , matl_landed_cost_php
		 , pi_fg_process_php
		 , pi_resin_php
		 , pi_hidden_profit_php
		 , sf_lbr_cost_php
		 , sf_ovhd_cost_php
		 , fg_lbr_cost_php
		 , fg_ovhd_cost_php
		 , total_actl_unit
		 , nolanded_actl_unit
	FROM #RMBreakDownFG
	WHERE [Level] <> 0

END