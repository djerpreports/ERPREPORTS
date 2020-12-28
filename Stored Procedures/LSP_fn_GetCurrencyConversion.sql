CREATE FUNCTION LSP_fn_GetCurrencyConversion (
--DECLARE  
	@TransDate			DATETIME		--= '05/01/2020' 
  , @FromCurrency		CurrCodeType	--= 'JPY'  
  , @ToCurrency			CurrCodeType	--= 'USD'  
  , @Amount				AmountType		--= '10000.00'  
  
) RETURNS AmountType  

BEGIN
  
	DECLARE  
		@FromIsDivisor		INT  
	  , @ToIsDivisor		INT  
	  , @ExchangeRate		ExchRateType 
	  , @ConvertedAmount	AmountType 
	  
	SET @TransDate = ISNULL(@TransDate, GETDATE())  
	SET @FromCurrency = ISNULL(@FromCurrency, 'PHP')  
	SET @ToCurrency = ISNULL(@ToCurrency, 'PHP')  
	SET @Amount = ISNULL(@Amount, 0)  
	SET @ConvertedAmount = ISNULL(@ConvertedAmount, 0)  
	SET @ExchangeRate = ISNULL(@ExchangeRate, 0)  
	  
	SET @FromIsDivisor = (SELECT ISNULL(rate_is_divisor, 0) FROM currency WHERE curr_code = @FromCurrency)  
	SET @ToIsDivisor = (SELECT ISNULL(rate_is_divisor, 0) FROM currency WHERE curr_code = @ToCurrency)  
	  
	IF @FromCurrency = @ToCurrency  
	BEGIN  
		SET @ExchangeRate = 1  
		SET @ConvertedAmount = @Amount * @ExchangeRate  
	END  
	ELSE  
		BEGIN  
		  
			SET @ExchangeRate =   
				(SELECT sell_rate FROM   
				  (SELECT TOP(1) ISNULL(sell_rate,0) AS sell_rate, ISNULL(eff_date,GETDATE()) AS eff_date, ISNULL(from_curr_code, '') AS curr_code  
				   FROM currate  
				   WHERE from_curr_code = @FromCurrency AND to_curr_code = @ToCurrency  
					 AND YEAR(eff_date) = YEAR(@TransDate)  
					 AND MONTH(eff_date) <= MONTH(@TransDate)  
				   ORDER BY eff_date DESC) AS curr1  
			UNION ALL  
			SELECT sell_rate FROM   
				  (SELECT TOP(1) ISNULL(sell_rate,0) AS sell_rate, ISNULL(eff_date, GETDATE()) AS eff_date, ISNULL(from_curr_code, '') AS curr_code  
				   FROM currate  
				   WHERE from_curr_code = @ToCurrency AND to_curr_code = @FromCurrency  
					 AND YEAR(eff_date) = YEAR(@TransDate)  
					 AND MONTH(eff_date) <= MONTH(@TransDate)  
				   ORDER BY eff_date DESC) AS curr2
				)  
			   
			 IF @ToIsDivisor = 1 AND @FromCurrency = 'PHP'  
			 BEGIN  
				 SET @ConvertedAmount = (@Amount / NULLIF(@ExchangeRate, 0))  
			 END  
				 ELSE IF @FromIsDivisor = 1 AND @ToCurrency = 'PHP'  
			 BEGIN  
				SET @ConvertedAmount = (@Amount * NULLIF(@ExchangeRate, 0))  
			 END  
			 ELSE IF @ToIsDivisor = 1 AND @FromCurrency <> 'PHP'  
			 BEGIN  
				 SET @ConvertedAmount = (@Amount * @ExchangeRate)  
			 END  
			 ELSE IF @FromIsDivisor = 1 AND @ToCurrency <> 'PHP'  
			 BEGIN  
				 SET @ConvertedAmount = (@Amount / @ExchangeRate)  
			 END  
				 ELSE IF @ToIsDivisor <> 1 AND @FromCurrency = 'PHP'  
			 BEGIN  
				 SET @ConvertedAmount = (@Amount * NULLIF(@ExchangeRate, 0))  
			 END  
			 ELSE IF @ToIsDivisor <> 1 AND @FromCurrency <> 'PHP'  
			 BEGIN  
				SET @ConvertedAmount = (@Amount / NULLIF(@ExchangeRate, 0))  
			 END  
			 ELSE  
			 BEGIN  
				 SET @ConvertedAmount = @Amount  
				 SET @ExchangeRate = 1  
			 END  
	END  
	
	RETURN @ConvertedAmount

END