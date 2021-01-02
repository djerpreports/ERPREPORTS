CREATE PROCEDURE LSP_Rpt_NewDM_InventoryTurnOverReportSP (
--DECLARE
	@IsShowDetail					BIT = 1
) AS
BEGIN

	SELECT *
	FROM Rpt_InventoryTurnOver

END


CREATE PROCEDURE LSP_Rpt_NewDM_MiscellaneousTransactionReportSp (
--DECLARE
	@StartDate					DateType	--= '05/01/2020'
  , @EndDate					DateType	--= '05/31/2020'
) AS
BEGIN

	SELECT *
	FROM Rpt_MiscTransaction

END