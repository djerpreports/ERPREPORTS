ALTER PROCEDURE LSP_StdCost_GetMatlCostingSp (
--DECLARE
	@Matl			ItemType		--= 'PI-FG-MDK3002'
  , @TransDate		DateType		--= '05/29/2020'
  , @MatlUnitCost	DECIMAL(18,10)	= 0
									OUTPUT
  , @PIProcessCost	DECIMAL(18,10)	= 0
  									OUTPUT
  , @PIResinCost	DECIMAL(18,10)	= 0
  									OUTPUT
  , @PIHiddenProfit	DECIMAL(18,10)	= 0
  									OUTPUT
  , @SFLbrCst		DECIMAL(18,10)	= 0
  									OUTPUT
  , @SFOvhdCst		DECIMAL(18,10)	= 0
  									OUTPUT
  , @FGLbrCst		DECIMAL(18,10)	= 0
  									OUTPUT
  , @FGOvhdCst		DECIMAL(18,10)	= 0
  									OUTPUT
) AS
BEGIN
	DECLARE 
		@PIVendCost		DECIMAL(18,10)
	  , @LaborRate		CostPrcType  
	  , @OvhdRate		CostPrcType  
	  , @LaborCost		AmountType  
	  , @OverhdCost		AmountType  
	  , @IsPIFG			INT
	
	IF @Matl NOT LIKE 'SF-%' AND @Matl NOT LIKE 'FG-%'
	BEGIN
		
		SELECT @IsPIFG = CASE WHEN COUNT(*) >= 1 THEN 1 ELSE 0 END
		FROM itemvend AS iv
				JOIN itemvendprice AS ivp 
					ON iv.item = ivp.item
					  AND iv.vend_num = ivp.vend_num
					  AND iv.[rank] = 1
		WHERE iv.item = @Matl
		  AND iv.vend_num = 'LPI0001'		  
		
		IF @IsPIFG = 1
		BEGIN
			SELECT @PIProcessCost = CASE WHEN iv.item LIKE 'PI-FG-%' OR (iv.item LIKE 'SC-%' AND ivp.vend_num = 'LPI0001')
											THEN ISNULL(ivp.Uf_process_cost, 0)
										 ELSE 0.00 END
				 , @PIResinCost = CASE WHEN iv.item LIKE 'PI-FG-%' OR (iv.item LIKE 'SC-%' AND ivp.vend_num = 'LPI0001')
											THEN ISNULL(ivp.Uf_resin_cost, 0)
										 ELSE 0.00 END
				 , @PIVendCost = CASE WHEN iv.item LIKE 'PI-FG-%' OR (iv.item LIKE 'SC-%' AND ivp.vend_num = 'LPI0001')
											THEN ISNULL(ivp.brk_cost##1, 0)
										 ELSE 0.00 END			 
			FROM itemvend AS iv
				JOIN itemvendprice AS ivp 
					ON iv.item = ivp.item
					  AND iv.vend_num = ivp.vend_num
					  AND iv.[rank] = 1
			WHERE iv.item = @Matl
			
			SET @PIHiddenProfit = @PIVendCost - (@PIProcessCost + @PIResinCost)
		END		
		ELSE	--IF @Matl LIKE 'RM-%' OR (@Matl LIKE 'SF-%' AND @IsPIFG = 0)
		BEGIN
		
			SELECT TOP(1) @MatlUnitCost = CASE WHEN ip.curr_code = 'USD'
													THEN unit_price1 / 1.2
												ELSE
													dbo.LSP_fn_GetCurrencyConversion(@TransDate, ip.curr_code, 'USD', ip.unit_price1)
										  END
			FROM item AS i
				JOIN itemprice AS ip
					ON i.item = ip.item
			WHERE i.item = @Matl
			  AND ip.effect_date <= @TransDate
			ORDER BY effect_date DESC
		
		END
		
		SELECT @SFLbrCst = 0
			 , @SFOvhdCst = 0
			 , @FGLbrCst = 0
			 , @FGOvhdCst = 0
	END
	ELSE
	BEGIN
		
		SELECT TOP(1)   
			@LaborRate = rm_labor_rate  
		  , @OvhdRate = rm_ovhd_rate  
		  
		FROM LSP_labor_oh_rate  
		WHERE effective_date <= GETDATE()  
		ORDER BY effective_date DESC  
		  
		SELECT @LaborCost = (SUM(js.run_lbr_hrs) * 60 * @LaborRate)
		FROM item AS i 
			JOIN jrt_sch AS js
				ON i.job = js.job  
				  AND i.suffix = js.suffix
		WHERE i.item = @Matl 
		         
		SET @OverhdCost = @LaborCost * @OvhdRate  
		  
		IF @Matl LIKE 'SF-%'
		BEGIN
			SELECT @SFLbrCst = @LaborCost
				 , @SFOvhdCst = @OverhdCost
		END
		ELSE IF @Matl LIKE 'FG-%'
		BEGIN
			SELECT @FGLbrCst = @LaborCost
				 , @FGOvhdCst = @OverhdCost
		END
	
	END

--SELECT @MatlUnitCost, @PIProcessCost, @PIResinCost, @PIHiddenProfit, @SFLbrCst, @SFOvhdCst, @FGLbrCst, @FGOvhdCst		

END