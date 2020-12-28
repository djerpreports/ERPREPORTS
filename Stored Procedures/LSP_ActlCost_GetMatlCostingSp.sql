ALTER PROCEDURE LSP_ActlCost_GetMatlCostingSp (
--DECLARE
	@matl_item					ItemType		--= 'SF-3DK3002'
  , @matl_lot					LotType			--= '19RF-00070'
  , @matlTransDate				DateType		--= '2020-02-13 10:44:25.000'
  , @JobQty						BIGINT			
										OUTPUT
  , @matl_unit_cost_usd     AmountType --= 0
                OUTPUT
  , @matl_landed_cost_usd     AmountType --= 0
                OUTPUT
  , @pi_fg_process_usd      AmountType --= 0
                OUTPUT
  , @pi_resin_usd     AmountType --= 0
                OUTPUT
  , @pi_vend_cost_usd     AmountType --= 0
                OUTPUT
  , @pi_hidden_profit_usd     AmountType --= 0
                OUTPUT
  , @sf_lbr_cost_usd      AmountType --= 0
                OUTPUT
  , @sf_ovhd_cost_usd     AmountType --= 0
                OUTPUT
  , @fg_lbr_cost_usd      AmountType --= 0
                OUTPUT
  , @fg_ovhd_cost_usd     AmountType --= 0
                OUTPUT
  , @matl_unit_cost_php     AmountType --= 0
                OUTPUT
  , @matl_landed_cost_php     AmountType --= 0
                OUTPUT
  , @pi_fg_process_php      AmountType --= 0
                OUTPUT
  , @pi_resin_php     AmountType --= 0
                OUTPUT
  , @pi_vend_cost_php     AmountType --= 0
                OUTPUT
  , @pi_hidden_profit_php     AmountType --= 0
                OUTPUT
  , @sf_lbr_cost_php      AmountType --= 0
                OUTPUT
  , @sf_ovhd_cost_php     AmountType --= 0
                OUTPUT
  , @fg_lbr_cost_php      AmountType --= 0
                OUTPUT
  , @fg_ovhd_cost_php     AmountType --= 0
                OUTPUT
) AS

BEGIN
	DECLARE 
		@pi_fg_process			AmountType
	  , @pi_resin				AmountType
	  , @ExchRate			ExchRateType
	  , @CurrCode			CurrCodeType
	  , @ReceiptCount		BIGINT		= 0
	  , @LaborRate		CostPrcType  
	  , @OvhdRate		CostPrcType  
	  , @LaborCost		AmountType  
	  , @OverhdCost		AmountType 
	
	SET @JobQty = 1
	
	/****FOR MATERIALS FROM PLASTIC INJECTION, vend_num = 'LPI0001' FROM ITEM VENDOR RANK 1 ****/
	IF EXISTS (SELECT * FROM itemvend WHERE item = @matl_item AND [rank] = 1 AND vend_num = 'LPI0001')
		  AND (@matl_item NOT LIKE 'FG-%' AND @matl_item NOT LIKE 'SF-%' )
	BEGIN
		SELECT @pi_fg_process = CASE WHEN iv.item LIKE 'PI-FG-%' OR (iv.item LIKE 'SC-%' AND ivp.vend_num = 'LPI0001')
										THEN ISNULL(ivp.Uf_process_cost, 0)
									 ELSE 0.00 END
			 , @pi_resin = CASE WHEN iv.item LIKE 'PI-FG-%' OR (iv.item LIKE 'SC-%' AND ivp.vend_num = 'LPI0001')
										THEN ISNULL(ivp.Uf_resin_cost, 0)
									 ELSE 0.00 END				 
			 , @matl_landed_cost_usd = 0
			 , @CurrCode = v.curr_code
		FROM itemvend AS iv
			JOIN itemvendprice AS ivp 
				ON iv.item = ivp.item
				  AND iv.vend_num = ivp.vend_num
				  AND iv.[rank] = 1
			JOIN vendor AS v
				ON iv.vend_num = v.vend_num
		WHERE iv.item = @matl_item
		
		
			EXEC dbo.LSP_CurrencyConversionModSp @matlTransDate, @CurrCode, 'USD', @pi_fg_process, @pi_fg_process_usd OUTPUT, @ExchRate OUTPUT
			EXEC dbo.LSP_CurrencyConversionModSp @matlTransDate, @CurrCode, 'USD', @pi_resin, @pi_resin_usd OUTPUT, @ExchRate OUTPUT			
			EXEC dbo.LSP_CurrencyConversionModSp @matlTransDate, @CurrCode, 'PHP', @pi_fg_process, @pi_fg_process_php OUTPUT, @ExchRate OUTPUT			
			EXEC dbo.LSP_CurrencyConversionModSp @matlTransDate, @CurrCode, 'PHP', @pi_resin, @pi_resin_php OUTPUT, @ExchRate OUTPUT			
		
		SELECT @pi_vend_cost_php = ISNULL(matl_cost, 0)
		FROM matltran AS m
		WHERE m.trans_type = 'R'
		  AND m.lot = @matl_lot
		  AND m.item = @matl_item			
		
		EXEC dbo.LSP_CurrencyConversionModSp @matlTransDate, 'PHP', 'USD', @pi_vend_cost_php, @pi_vend_cost_usd OUTPUT, @ExchRate OUTPUT			
		
		SET @pi_hidden_profit_php = @pi_vend_cost_php - (@pi_fg_process_php + @pi_resin_php)
		SET @pi_hidden_profit_usd = @pi_vend_cost_usd - (@pi_fg_process_usd + @pi_resin_usd)
		
		--SELECT @pi_vend_cost_usd, @matl_item, @matl_lot, @pi_hidden_profit_usd
		
	END	
	/****FOR MATERIALS NOT FROM PLASTIC INJECTION AND NOT SF OR FG ****/
	ELSE IF @matl_item NOT LIKE 'FG-%' AND @matl_item NOT LIKE 'SF-%'
		BEGIN
		
			SELECT @CurrCode = MAX(v.curr_code)
				 , @ReceiptCount = COUNT(*)
			FROM matltran AS m
				JOIN poitem AS poi 
					ON m.ref_num = poi.po_num AND m.ref_line_suf = poi.po_line 
				JOIN po 
					ON poi.po_num = po.po_num 
				JOIN vendor AS v
					ON po.vend_num = v.vend_num
			WHERE m.item = @matl_item
			  AND m.lot = @matl_lot
			  AND m.trans_type = 'R'
		
			IF @ReceiptCount = 0
			BEGIN
			
				IF (@matl_item LIKE 'PDN-%')
				BEGIN
					SELECT TOP(1) @matl_unit_cost_php = matl_cost
					FROM matltran
					WHERE trans_type = 'H' AND item = @matl_item AND lot = @matl_lot
					
					EXEC dbo.LSP_CurrencyConversionModSp @matlTransDate, 'PHP', 'USD', @matl_unit_cost_php, @matl_unit_cost_usd OUTPUT, @ExchRate OUTPUT
				END
				ELSE
				BEGIN						
				
					SELECT TOP(1) @matl_unit_cost_usd = (unit_price1 / 1.2)
					FROM itemprice
					WHERE item = @matl_item
					  AND effect_date <= @matlTransDate
					ORDER BY effect_date DESC
					
					EXEC dbo.LSP_CurrencyConversionModSp @matlTransDate, 'USD', 'PHP', @matl_unit_cost_usd, @matl_unit_cost_php OUTPUT, @ExchRate OUTPUT
					
				END
				
				SELECT @matl_landed_cost_usd = 0
					 , @matl_landed_cost_php = 0
			END
			ELSE
			BEGIN
				SELECT TOP(1) @matl_unit_cost_php = (por.unit_mat_cost * por.exch_rate)
					 , @matl_landed_cost_php = (por.unit_duty_cost + por.unit_brokerage_cost + por.unit_freight_cost + por.unit_loc_frt_cost) * por.exch_rate				
				FROM matltran AS m
					JOIN po_rcpt AS por
						ON m.ref_num = por.po_num 
						  AND m.ref_line_suf = por.po_line 
						  AND m.trans_date = por.rcvd_date
				WHERE m.item = @matl_item 
				  AND m.lot = @matl_lot 
				  AND m.trans_type = 'R'
				  
				EXEC dbo.LSP_CurrencyConversionModSp @matlTransDate, 'PHP', 'USD', @matl_unit_cost_php, @matl_unit_cost_usd OUTPUT, @ExchRate OUTPUT
				EXEC dbo.LSP_CurrencyConversionModSp @matlTransDate, 'PHP', 'USD', @matl_landed_cost_php, @matl_landed_cost_usd OUTPUT, @ExchRate OUTPUT

			END			
			
		END
	/****FOR ISSUED JOB MATERIALS WHERE ITEM IS FG ****/
	ELSE IF @matl_item LIKE 'FG-%'
	BEGIN
	
		IF EXISTS (SELECT * FROM job WHERE job = @matl_lot)
		BEGIN
			SELECT TOP(1)   
				@LaborRate = rm_labor_rate  
			  , @OvhdRate = rm_ovhd_rate  
			  
			FROM LSP_labor_oh_rate  
			WHERE effective_date <= GETDATE()  
			ORDER BY effective_date DESC  
			  
			SELECT @LaborCost = (SUM(js.run_lbr_hrs) * 60 * @LaborRate)
			FROM jrt_sch AS js
			WHERE js.job = @matl_lot
			  AND js.suffix = 0
			--FROM item AS i 
			--	JOIN jrt_sch AS js
			--		ON i.job = js.job  
			--		  AND i.suffix = js.suffix
			--WHERE i.item = @matl_item			  
			         
			SET @OverhdCost = @LaborCost * @OvhdRate
			
			SELECT @fg_lbr_cost_php = @LaborCost
				 , @fg_ovhd_cost_php = @OverhdCost
			
			SELECT @jobQty = qty_released
			FROM job
			WHERE job = @matl_lot
			  AND suffix = 0
			
			EXEC dbo.LSP_CurrencyConversionModSp @matlTransDate, 'PHP', 'USD', @fg_lbr_cost_php, @fg_lbr_cost_usd OUTPUT, @ExchRate OUTPUT
			EXEC dbo.LSP_CurrencyConversionModSp @matlTransDate, 'PHP', 'USD', @fg_ovhd_cost_php, @fg_ovhd_cost_usd OUTPUT, @ExchRate OUTPUT	
			
	
		END
		ELSE IF EXISTS(SELECT * FROM matltran WHERE lot = @matl_lot AND trans_type = 'R' AND item = @matl_item)
		BEGIN			
		
			SELECT @matl_unit_cost_php = matl_cost
			FROM matltran 
			WHERE lot = @matl_lot 
			  AND trans_type = 'R' 
			  AND item = @matl_item
			  
			EXEC dbo.LSP_CurrencyConversionModSp @matlTransDate, 'PHP', 'USD', @matl_unit_cost_php, @matl_unit_cost_usd OUTPUT, @ExchRate OUTPUT
			
		END
		ELSE
		BEGIN
			SELECT TOP(1) @matl_unit_cost_usd = (unit_price1 / 1.2)
			FROM itemprice
			WHERE item = @matl_item
			  AND effect_date <= @matlTransDate
			ORDER BY effect_date DESC
				
			EXEC dbo.LSP_CurrencyConversionModSp @matlTransDate, 'USD', 'PHP', @matl_unit_cost_usd, @matl_unit_cost_php OUTPUT, @ExchRate OUTPUT					
			
		END
	END
	/****FOR ISSUED JOB MATERIALS WHERE ITEM IS SF ****/
	ELSE IF @matl_item LIKE 'SF-%'
	BEGIN
	
		IF EXISTS (SELECT * FROM job WHERE job = @matl_lot)
		BEGIN
			SELECT TOP(1)   
				@LaborRate = rm_labor_rate  
			  , @OvhdRate = rm_ovhd_rate  
			  
			FROM LSP_labor_oh_rate  
			WHERE effective_date <= GETDATE()  
			ORDER BY effective_date DESC  
			  
			SELECT @LaborCost = (SUM(js.run_lbr_hrs) * 60 * @LaborRate)
			FROM jrt_sch AS js
			WHERE js.job = @matl_lot
			--FROM item AS i 
			--	JOIN jrt_sch AS js
			--		ON i.job = js.job  
			--		  AND i.suffix = js.suffix
			--WHERE i.item = @matl_item 
			         
			SET @OverhdCost = @LaborCost * @OvhdRate
			
			SELECT @sf_lbr_cost_php = @LaborCost
				 , @sf_ovhd_cost_php = @OverhdCost
			
			SELECT @jobQty = qty_released
			FROM job
			WHERE job = @matl_lot
			  AND suffix = 0
			
			EXEC dbo.LSP_CurrencyConversionModSp @matlTransDate, 'PHP', 'USD', @sf_lbr_cost_php, @sf_lbr_cost_usd OUTPUT, @ExchRate OUTPUT
			EXEC dbo.LSP_CurrencyConversionModSp @matlTransDate, 'PHP', 'USD', @sf_ovhd_cost_php, @sf_ovhd_cost_usd OUTPUT, @ExchRate OUTPUT			
	
		END		
		ELSE IF EXISTS(SELECT * FROM matltran WHERE lot = @matl_lot AND trans_type = 'H' AND item = @matl_item)
		BEGIN			
		
			SELECT TOP(1) @matl_unit_cost_php = matl_cost
			FROM matltran 
			WHERE lot = @matl_lot 
			  AND trans_type = 'H' 
			  AND item = @matl_item
			  AND trans_date <= @matlTransDate
			ORDER BY trans_date DESC
			  
			EXEC dbo.LSP_CurrencyConversionModSp @matlTransDate, 'PHP', 'USD', @matl_unit_cost_php, @matl_unit_cost_usd OUTPUT, @ExchRate OUTPUT
			
		END
		ELSE
		BEGIN
			SELECT TOP(1) @matl_unit_cost_usd = (unit_price1 / 1.2)
			FROM itemprice
			WHERE item = @matl_item
			  AND effect_date <= @matlTransDate
			ORDER BY effect_date DESC
				
			EXEC dbo.LSP_CurrencyConversionModSp @matlTransDate, 'USD', 'PHP', @matl_unit_cost_usd, @matl_unit_cost_php OUTPUT, @ExchRate OUTPUT					
			
		END
	END
	
--SELECT @JobQty, @matl_unit_cost_usd AS matl_unit_cost_usd
--  , @matl_landed_cost_usd AS matl_landed_cost_usd
--  , @pi_fg_process_usd  AS pi_fg_process_usd
--  , @pi_resin_usd AS pi_resin_usd
--  , @pi_vend_cost_usd AS pi_vend_cost_usd
--  , @pi_hidden_profit_usd AS pi_hidden_profit_usd
--  , @sf_lbr_cost_usd  AS sf_lbr_cost_usd
--  , @sf_ovhd_cost_usd AS sf_ovhd_cost_usd
--  , @fg_lbr_cost_usd  AS fg_lbr_cost_usd
--  , @fg_ovhd_cost_usd AS fg_ovhd_cost_usd
--  , @matl_unit_cost_php AS matl_unit_cost_php
--  , @matl_landed_cost_php AS matl_landed_cost_php
--  , @pi_fg_process_php  AS pi_fg_process_php
--  , @pi_resin_php AS pi_resin_php
--  , @pi_vend_cost_php AS pi_vend_cost_php
--  , @pi_hidden_profit_php AS pi_hidden_profit_php
--  , @sf_lbr_cost_php  AS sf_lbr_cost_php
--  , @sf_ovhd_cost_php AS sf_ovhd_cost_php
--  , @fg_lbr_cost_php  AS fg_lbr_cost_php
--  , @fg_ovhd_cost_php AS fg_ovhd_cost_php

END