SELECT
  c.ID AS ClothItemID,
  c.ProductionID,
  c.Item AS ClothItem,
  c.Description,
  c.Qty AS ClothQty,
  c.Status AS ClothStatus,
  c.Date AS ClothDate,
  prod.Inventory AS InventoryID,
  inventory.Item AS InventoryItem,
  inventory.ItemClass,
  CONCAT("[", FORMAT_DATE('%d %b %y', inventory.Date),"]-","[",inventory.ReceivingBarcode,"]-[",inventory.ItemDescription,"]-[",inventory.Color,"]-[",inventory.KGs_Mtrs_PCs,inventory.Metric,"]") AS InventoryDetails,
  inventory.ItemDescription,
  inventory.Color,
  inventory.ReceivingBarcode,
  inventory.Status AS InventoryStatus,
  inventory.Metric,
  Prod.Status,
  c.PatternsDamages,
  c.OperationsDamages,
  c.QcDamages,
  c.Qty - COALESCE(c.PatternsDamages, 0) - COALESCE(c.OperationsDamages, 0) - COALESCE(c.QcDamages, 0) AS FinishedClothes,
  COALESCE(c.PatternsDamages, 0) + COALESCE(c.OperationsDamages, 0) + COALESCE(c.QcDamages, 0) AS TotalDamages
FROM
  `appsheetprojects.VuelaProduction.ClothItems` AS c
LEFT JOIN (
  SELECT
    P.ID,
    P.Inventory,
    P.Date,
    P.IssueTo,
    P.Issued_Kgs_Mtrs_Pcs_,
    P.DateissuedtoVintagestore,
    P.QtyIssued_Vintage_,
    P.Status
  FROM
    `VuelaProduction.Production` AS P ) AS prod
ON
  c.ProductionID = prod.ID
LEFT JOIN (
  SELECT
    s.ID,
    s.Date,
    s.Item,
    s.ItemClass,
    s.ItemDescription,
    s.Color,
    s.ReceivingBarCode,
    s.Metric,
    s.KGs_Mtrs_PCs,
    s.Status,
    s.UnitCost
  FROM
    `VuelaProduction.Store` AS s ) AS inventory
ON
  inventory.ID = prod.inventory
WHERE
  c.Status = 'Active'
  AND inventory.Status <>'Cancelled'	
