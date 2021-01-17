ALTER PROCEDURE LSP_Rpt_NewDM_MiscellaneousTransactionReportSp (
--DECLARE
	@StartDate					DateType	--= '05/01/2020'
  , @EndDate					DateType	--= '05/31/2020'
) AS
BEGIN

	SELECT TOP(20) WITH TIES *
	FROM Rpt_MiscTransaction
	ORDER BY ROW_NUMBER() OVER (PARTITION BY TransDesc, Wc ORDER BY TransDate)
END