ALTER PROCEDURE LSP_Rpt_ProductivityIndex_DetailedSp (  
--DECLARE  
	@MonthYear		NVARCHAR(20) --= 'Feb-2020'
  , @IsSummary		BIT		--= 0
) AS  
  
BEGIN
	DECLARE
		@StartDate		DateType --= '03/01/2020'  
	  , @EndDate		DateType --= '03/31/2020'  
  

	IF OBJECT_ID('tempdb..##productivity') IS NOT NULL
		DROP TABLE ##productivity

	SET @StartDate = CAST( ('01-' + @MonthYear) AS DATETIME)	
	SET @EndDate = DATEADD(S, -1, DATEADD(month, 1, @StartDate))
    
    
	SELECT jt.job  
		  , jt.suffix  
		  , j.item  
		  , i.description  AS item_desc
		  , jt.trans_date  
		  , jt.trans_type  
		  , jt.qty_complete  
		  , jt.oper_num  
		  , jt.wc  
		  , w.description AS wc_desc
		  , js.run_lbr_hrs  AS std_unit
		  , js.run_lbr_hrs * 60 AS labor_min_per_pc
		  , (js.run_lbr_hrs * jt.qty_complete)  AS std_labor_hrs
		  , ISNULL(jt.a_hrs,0)  AS actl_labor_hrs
		  , (js.run_lbr_hrs * jt.qty_complete)  * 60 AS std_labor_min
		  , ISNULL(jt.a_hrs,0) * 60 AS actl_labor_min
		  , LTRIM(RTRIM(ISNULL(w.Uf_workcenter_class, ''))) AS WorkCenterClassification
		  , ISNULL(jt.start_time , 0) AS start_time
		  , ISNULL(jt.end_time  , '') AS end_time
	INTO ##productivity
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
		JOIN  wc AS w 
		ON jt.wc = w.wc  
       
	WHERE (jt.trans_type = 'R' OR jt.trans_type = 'M') AND jt.trans_date >= @StartDate AND jt.trans_date <= @EndDate  
	 
	  --AND jt.wc <> 'FG' AND jt.wc <> 'PAINSP' AND jt.wc <> 'PI-CUT' AND jt.wc <> 'PI-MLD' AND jt.wc <> 'PI-PCK'  
	  --AND jt.wc <> 'PRINT' AND jt.wc <> 'SC-INS'  
  
	
  
	--UPDATE THE REPORT TABLE. CONVERT HOURS TO MINUTES.  
	--UPDATE @report_table  
	--SET std_labor_min = std_labor_hrs * 60  
	--  , actl_labor_min = actl_labor_hrs * 60  
      
	IF @IsSummary = 0
	BEGIN
		SELECT * 
		FROM ##productivity
		--WHERE WOrkCenterClassification LIKE '%pack%'
		ORDER BY job, trans_date, oper_num  
		
		DROP TABLE ##productivity
	END
	

END