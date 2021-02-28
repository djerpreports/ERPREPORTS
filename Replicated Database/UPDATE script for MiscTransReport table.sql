SELECT * 
FROM Rpt_MiscTransaction 
WHERE MiscTransClass = '' AND ReasonDesc = 'Section Requests'

--UPDATE Rpt_MiscTransaction 
--SET TransQty = 1
--  , MatlCost_PHP = -81911.11128283
--  , MatlLandedCost_PHP = -94.00154739
--  , PIFGProcess_PHP = -1563.54929012
--  , PIResin_PHP = -1161.05922924
--  , PIHiddenProfit_PHP = -1585.94694603
--  , SFAddedCost_PHP = -21.86240045
--  , TotalCost_PHP = -86337.53069606488
--WHERE MiscTransClass = '' AND ReasonDesc = 'SCRAP'

UPDATE Rpt_MiscTransaction 
SET TransQty = 1
  , MatlCost_PHP = -7719.1371856200000000
  , MatlLandedCost_PHP = -505.7795999500000000
  , PIFGProcess_PHP = -325.1525520000000000
  , PIResin_PHP = -120.7829088000000000
  , PIHiddenProfit_PHP = -416.1569392000000000
  , SFAddedCost_PHP = -335.1712343000000000
  , TotalCost_PHP = -9422.18041987000
WHERE MiscTransClass = '' AND ReasonDesc = 'Section Requests'
