ALTER PROCEDURE LSP_Rpt_NewDM_InventoryTurnOverReportSP (
--DECLARE
	@IsShowDetail					BIT = 1
  , @StartDate					DATETIME		OUTPUT
  , @EndDate					DATETIME		OUTPUT
) AS
BEGIN

	SELECT @StartDate =	'04/01/2020'
		 , @EndDate = '03/31/2021'

	SELECT * 
	FROM Rpt_InvtyTurnover2

END