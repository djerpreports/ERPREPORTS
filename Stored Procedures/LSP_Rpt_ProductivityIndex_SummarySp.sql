--ALTER PROCEDURE LSP_Rpt_ProductivityIndex_SummarySp (
DECLARE
	@MonthYear		NVARCHAR(20) = 'Mar-2020'
  , @IsSummary		BIT		= 1
--) AS  
  
BEGIN
	DECLARE
		@StartDate		DateType 
	  , @EndDate		DateType 
	  , @MidDate		DateType

	SET @StartDate = CAST( ('01-' + @MonthYear) AS DATETIME)	
	SET @EndDate = DATEADD(S, -1, DATEADD(month, 1, @StartDate))

	SELECT @MidDate = CAST(MONTH(@StartDate) AS NVARCHAR(2))+ '/15/' + CAST(YEAR(@StartDate) AS NVARCHAR(4))

	EXEC dbo.LSP_Rpt_ProductivityIndex_DetailedSp @MonthYear, 1;

	WITH CTE_wc AS (
	SELECT DISTINCT LTRIM(RTRIM(Uf_workcenter_class)) AS WorkCenterClassification
	FROM wc
	WHERE Uf_workcenter_class IS NOT NULL
	UNION
	SELECT 'MH' )
	SELECT COALESCE(wc.WorkCenterClassification, rpt1.WorkCenterClassification, rpt2.WorkCenterClassification, rpt3.WorkCenterClassification) AS WorkCenterClassification
		-- , rpt1.WorkCenterClassification
		 , ISNULL(rpt1.std_labor_min, 0) AS std_labor_min_1st
		 , ISNULL(rpt1.actl_labor_min, 0) AS actl_labor_min_1st
		 , ISNULL(rpt2.std_labor_min, 0) AS std_labor_min_2nd
		 , ISNULL(rpt2.actl_labor_min, 0) AS actl_labor_min_2nd
		 , ISNULL(rpt3.std_labor_min, 0) AS std_labor_min_all
		 , ISNULL(rpt3.actl_labor_min, 0) AS actl_labor_min_all
	FROM CTE_wc AS wc
		FULL OUTER JOIN
				(SELECT WorkCenterClassification
					 , SUM(std_labor_min) AS std_labor_min
					 , SUM(actl_labor_min) AS actl_labor_min
				FROM ##productivity	
				GROUP BY WorkCenterClassification) AS rpt3
			ON wc.WorkCenterClassification = rpt3.WorkCenterClassification
		LEFT OUTER JOIN
				(SELECT WorkCenterClassification
					 , SUM(std_labor_min) AS std_labor_min
					 , SUM(actl_labor_min) AS actl_labor_min
				FROM ##productivity
				WHERE trans_date BETWEEN @StartDate AND dbo.DayEndOf(@MidDate)
				GROUP BY WorkCenterClassification) as rpt1
			ON rpt3.WorkCenterClassification = rpt1.WorkCenterClassification
		LEFT OUTER JOIN 
				(SELECT WorkCenterClassification
					 , SUM(std_labor_min) AS std_labor_min
					 , SUM(actl_labor_min) AS actl_labor_min
				FROM ##productivity
				WHERE trans_date BETWEEN DATEADD(Day,1,@MidDate) AND dbo.DayEndOf(@EndDate)
				GROUP BY WorkCenterClassification
				) AS rpt2
			ON rpt3.WorkCenterClassification = rpt2.WorkCenterClassification		
	ORDER BY wc.WorkCenterClassification;
	
	DROP TABLE ##productivity
END