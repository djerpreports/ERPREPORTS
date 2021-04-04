
CREATE PROCEDURE LSP_Rpt_NewDM_InventoryTurnOverReportSP (
	@IsShowDetail					BIT --= 1
) AS
BEGIN

	IF @IsShowDetail = 1
	BEGIN
		SELECT * 
		FROM [Rpt_InventoryTurnOver]
		--FROM Rpt_InvtyTurnOver
	END
	ELSE
	BEGIN
		SELECT * 
		FROM [Rpt_InventoryTurnOver]
		--FROM Rpt_InvtyTurnOver
		WHERE report_group <> 'DETAILED'
	END

END