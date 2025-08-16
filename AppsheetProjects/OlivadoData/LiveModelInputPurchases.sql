SELECT 
  Date_Purchased,
  Purchase_Price, 
  Amount, 
  Final_Price, 
  Purchase_Status,
  Amount / Final_Price AS Quantity
FROM `olivadofield.OlivadoFieldAnalytics.Purchases`
WHERE Purchase_Status IN ('Paid', 'Pending Payment')
  AND Date_Purchased > DATE('2024-12-31')
  AND Purchase_Price=18;
