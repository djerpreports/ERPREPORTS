CREATE PROCEDURE LSP_Rpt_NewDM_RMBreakdownSalesSp (
--DECLARE
	@StartDate				DateType --= '05/01/2020'
  , @EndDate				DateType --= '05/31/2020'
) AS
BEGIN

	IF OBJECT_ID('tempdb..#RMBreakDownSales') IS NOT NULL
		DROP TABLE #RMBreakDownSales

	DECLARE @ShipTrans AS TABLE (  
		TransDate				DateType
	  , Item					ItemType
	  , QtyShipped				QtyUnitType
	  , JobOrder				JobType
	  , JobSuffix				SuffixType
	  , PONumber				NVARCHAR(20)
	  
	)
	
	CREATE TABLE #RMBreakDownSales (
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
	  , @QtyShipped				INT
	  , @JobOrder				JobType
	  , @JobSuffix				SuffixType
	  , @PONum					NVARCHAR(20)
	  , @SQLStr					NVARCHAR(1000)
	  
	INSERT INTO @ShipTrans
	SELECT m.trans_date
	  , m.item
	  , m.qty
	  , (SELECT TOP(1) m2.ref_num
		 FROM matltran AS m2
		 WHERE m.lot = m2.lot AND m2.trans_type = 'F' AND m.item = m2.item
		 ORDER BY m2.trans_date DESC)
	  , (SELECT TOP(1) m2.ref_line_suf
		 FROM matltran AS m2
		 WHERE m.lot = m2.lot AND m2.trans_type = 'F' AND m.item = m2.item
		 ORDER BY m2.trans_date DESC)
	  , coi.Uf_ponum	  
	FROM matltran AS m
		JOIN item AS i
			ON m.item = i.item
		JOIN coitem AS coi
			ON m.ref_num = coi.co_num AND m.ref_line_suf = coi.co_line
		JOIN co AS c
			ON coi.co_num = c.co_num
		JOIN custaddr AS ca
			ON c.cust_num = ca.cust_num AND c.cust_seq = ca.cust_seq
		JOIN do_seq AS dos
			ON coi.co_num = dos.ref_num AND coi.co_line  = dos.ref_line
		JOIN do_hdr AS doh
			ON dos.do_num = doh.do_num
	WHERE m.trans_date BETWEEN dbo.MidnightOf(@StartDate) AND dbo.DayEndOf(@EndDate)
	  AND m.trans_type = 'S' AND m.ref_type = 'O' AND doh.stat = 'A'
	  
	GROUP BY m.trans_date, m.item, m.qty, coi.Uf_ponum, m.lot	  
	
	UNION
	
	SELECT m.trans_date
		 , m.item
		 , m.qty
		 , (SELECT TOP(1) m2.ref_num
			FROM matltran AS m2
			WHERE m.lot = m2.lot AND m2.trans_type = 'F' AND m.item = m2.item
			ORDER BY m2.trans_date DESC)
		 , (SELECT TOP(1) m2.ref_line_suf
			FROM matltran AS m2
			WHERE m.lot = m2.lot AND m2.trans_type = 'F' AND m.item = m2.item
			ORDER BY m2.trans_date DESC)
		 , coi.Uf_ponum		 
	FROM matltran AS m
		JOIN item AS i
			ON m.item = i.item
		LEFT OUTER JOIN coitem AS coi
			ON m.ref_num = coi.co_num AND m.ref_line_suf = coi.co_line
		JOIN co AS c
			ON coi.co_num = c.co_num
		JOIN custaddr AS ca
			ON c.cust_num = ca.cust_num AND c.cust_seq = ca.cust_seq
	
	WHERE m.trans_date BETWEEN dbo.MidnightOf(@StartDate) AND dbo.DayEndOf(@EndDate)
	  AND m.trans_type = 'S' AND m.ref_type = 'O' AND ca.name LIKE 'LSP%SECTION'
	GROUP BY m.trans_date, m.item, m.qty, coi.Uf_ponum, m.lot	  
		  
	DECLARE SalesCrsr CURSOR FAST_FORWARD FOR
	SELECT MAX(TransDate) TransDate
	  , PONumber
	  , JobOrder
	  , JobSuffix
	  , Item
	  , (SUM(QtyShipped) * (-1)) QtyShipped	
	FROM @ShipTrans
	GROUP BY PONumber,Item, JobOrder, JobSuffix
	HAVING (SUM(QtyShipped) * (-1)) <> 0
	ORDER BY PONumber

	--SELECT * 
	--FROM @FGReceipts

	OPEN SalesCrsr
	FETCH FROM SalesCrsr INTO
		@TransDate
	  , @PONum
	  , @JobOrder
	  , @JobSuffix
	  , @Item
	  , @QtyShipped
	
	WHILE @@FETCH_STATUS = 0
	BEGIN
		
				
		SET @SQLStr = 'SELECT * 
						FROM OPENROWSET(''SQLNCLI'', ''Server=SYTELINEERP803;Database=LSPI803_App;UID=sa;Pwd=P@$$w0rd'',
						''SET FMTONLY OFF;SET NOCOUNT ON; EXEC dbo.LSP_Rpt_NewDM_RMBreakdownPerJOSp ''''' + @JobOrder + ''''', ''''' + @PONum + ''''', '+ CAST(@QtyShipped AS NVARCHAR(10))  + ' '') AS a'
		PRINT @JobOrder + '_' +CAST(@QtyShipped AS NVARCHAR(10))

		INSERT INTO #RMBreakDownSales ( 
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

		UPDATE #RMBreakDownSales
		SET Item = @Item
		WHERE JONum = @JobOrder
		  AND PONum = @PONum

		FETCH NEXT FROM SalesCrsr INTO
			@TransDate
		  , @PONum
		  , @JobOrder
		  , @JobSuffix
		  , @Item
		  , @QtyShipped
	END
	
	CLOSE SalesCrsr
	DEALLOCATE SalesCrsr
	
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

	FROM #RMBreakDownSales
	WHERE [Level] <> 0
	
END