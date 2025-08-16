SELECT * ,
Amount / Final_Price AS QuantityPurchased
FROM `olivadofield.OlivadoFieldAnalytics.Purchases`
WHERE Purchase_Status IN ('Paid', 'Pending Payment','Pending Field Manager');
