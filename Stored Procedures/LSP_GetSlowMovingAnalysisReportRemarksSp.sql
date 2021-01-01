CREATE PROCEDURE LSP_GetSlowMovingAnalysisReportRemarksSp (
--DECLARE
	@Material					ItemType		--= 'PI-FG-MRT021'--'PI-FG-MMU108'
  , @Remarks					NVARCHAR(50)	OUTPUT
) AS

BEGIN

	IF OBJECT_ID('tempdb..#Level1Parent')
	 IS NOT NULL
		DROP TABLE #Level1Parent
	
	DECLARE
		@FGParentCount			INT
	  , @SFParentCount			INT
	  , @SFSubParentCount		INT

	SELECT i.item AS ParentItem
		 , jm.job  AS JobRef
		 , jm.item AS Material
	INTO #Level1Parent
	FROM jobmatl AS jm
		JOIN item AS i
			ON jm.job = i.job 
		JOIN job AS j
			ON jm.job = j.job AND jm.suffix = j.suffix  
	WHERE jm.item = @Material AND j.[type] = 'S'

	SELECT @FGParentCount = COUNT(*) FROM #Level1Parent WHERE ParentItem LIKE 'FG-%'
	SELECT @SFParentCount = COUNT(*) FROM #Level1Parent WHERE ParentItem LIKE 'SF-%'

	SELECT @SFSubParentCount = COUNT(*)
	FROM jobmatl AS jm
		JOIN item AS i
			ON jm.job = i.job 
		JOIN job AS j
			ON jm.job = j.job AND jm.suffix = j.suffix  
	WHERE jm.item IN (SELECT ParentItem 
					  FROM #Level1Parent
					  WHERE ParentItem LIKE 'SF-%')
	  AND j.[type] = 'S'
	  AND i.item LIKE 'FG-%'	  
	  
	SELECT @Remarks = CASE WHEN @FGParentCount = 0 AND @SFParentCount = 0
								THEN 'Not being used by any FG'
						   WHEN @SFParentCount > 0 AND @SFSubParentCount = 0
								THEN 'Being used by SF but SF is not being used by any FG'
						   ELSE ''
					  END

--SELECT * FROM #Level1Parent

	--SELECT @FGParentCount, @SFParentCount, @SFSubParentCount, @Remarks

END