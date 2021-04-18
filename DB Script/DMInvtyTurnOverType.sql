-- ================================
-- Create User-defined Table Type
-- ================================
USE LSPI803_App
GO

-- Create the data type
CREATE TYPE dbo.DMInvtyTurnOverType AS TABLE 
(
		trans_date					DATETIME  
	  , usage_matl					DECIMAL(19,8)  
	  , usage_landed				DECIMAL(19,8)
	  , product_code				NVARCHAR(100)
	  , invty_matl_cost				DECIMAL(19,8)
	  , invty_landed_cost			DECIMAL(19,8)
	  , safety_matl_cost			DECIMAL(19,8)
	  , report_group				NVARCHAR(10) 			
)
GO
