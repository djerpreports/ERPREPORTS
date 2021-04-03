CREATE PROCEDURE LSP_Rpt_NewDM_SalesSummaryReportSp (  
--DECLARE
	@StartDate					DateType	--= '05/01/2020'
  , @EndDate					DateType	--= '05/31/2020'
) AS  
  
BEGIN
  
	SELECT @StartDate = ISNULL(@StartDate, GETDATE())
		 , @EndDate = ISNULL(@EndDate, GETDATE())
  
	DECLARE @report_set AS TABLE (  
		inv_num   NVARCHAR(6)  
	  , sv_num   InvNumType  
	  , apply_to_inv InvNumType  
	  , inv_date  DateType  
	  , customer  NameType  
	  , cust_seq  CustSeqType  
	  , ship_to_cust NameType  
	  , inv_desc  DescriptionType  
	  , amount   AmountType  
	  , amount_php  AmountType  
	  , eng_design  AmountType  
	  , exch_rate  ExchRateType  
	  , price   AmountType  
	  , ar_desc   DescriptionType  
	  , do_num   DoNumType  
	  , type   ArTranTypeType  
	  , gl_ref   ReferenceType  	  
	)
	  
	INSERT INTO @report_set (  
	 inv_num  
	  , sv_num  
	  , apply_to_inv  
	  , inv_date  
	  , customer  
	  , cust_seq  
	  , ship_to_cust  
	  , inv_desc  
	  , amount  
	  , exch_rate  
	  , price  
	  , ar_desc  
	  , do_num  
	  , type  
	  , amount_php  
	  , eng_design  
	  , gl_ref  
	)  
	SELECT RIGHT(ar.description, 6)  
	  , ar.inv_num  
	  , ar.apply_to_inv_num  
	  , ar.inv_date
	  , ca2.name  
	  , do.cust_seq  
	  , ISNULL(ca.name, ' - ')  
	  , CASE ar.type  
	   WHEN 'C' THEN 'CREDIT MEMO'  
	   WHEN 'D' THEN 'DEBIT MEMO'  
	   ELSE CASE WHEN ar.do_num IS NULL  
		 THEN 'PRODUCT ENG'' DESIGN'  
		 ELSE 'PRODUCTS' END  
	   END  
	  , CASE WHEN ISNULL(inv.price, 0) < 0   
		THEN ar.amount * (-1)  
	   ELSE ar.amount END  
	  , ar.exch_rate  
	  , inv.price  
	  , ar.description  
	  , ar.do_num  
	  , ar.type  
	  , CASE WHEN ISNULL(inv.price, 0) < 0   
				THEN (ar.amount * ar.exch_rate) * (-1)  
			 ELSE (ar.amount * ar.exch_rate) END  
	  , CASE WHEN ar.do_num IS NULL AND ar.type = 'I'  
				THEN ar.amount  
			 ELSE 0.0 END  
	  , ar.ref  
	FROM artran AS ar
		INNER JOIN inv_hdr AS inv
			ON ar.inv_num = inv.inv_num
		LEFT OUTER JOIN do_hdr AS do
			ON ar.do_num = do.do_num
		LEFT OUTER JOIN custaddr AS ca
			ON do.cust_num = ca.cust_num AND do.cust_seq = ca.cust_seq
		LEFT OUTER JOIN custaddr AS ca2
			ON ar.cust_num = ca2.cust_num AND ca2.cust_seq = 0
	WHERE (ar.type <> 'P') AND ar.inv_date BETWEEN dbo.MidnightOf(@StartDate) AND dbo.DayEndOf(@EndDate)
	  
	--SELECT * FROM @gl_table  
	--SELECT * FROM @report_set
	SELECT inv_date
		 , inv_num
		 , ship_to_cust
		 , inv_desc
		 , amount
		 , amount_php
		 , exch_rate
		 , eng_design
		 , price
	FROM @report_set 
	ORDER BY inv_date
  
END