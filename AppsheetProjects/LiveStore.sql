
SELECT 
    s.*,
    s.UnitCost * s.KGs_Mtrs_PCs AS TotalCost,
    
    COALESCE(prod.IssuedUnits, 0) AS IssuedUnits,
    COALESCE(prod.IssuedUnits, 0) * s.UnitCost AS IssuedInventory,
    
    COALESCE(ret.ReturnedUnits, 0) AS ReturnedUnits,
    COALESCE(ret.ReturnedUnits, 0) * s.UnitCost AS ReturnedInventory,
    
    s.KGs_Mtrs_PCs - COALESCE(prod.IssuedUnits, 0) + COALESCE(ret.ReturnedUnits, 0) AS RemainingUnits,
    (s.KGs_Mtrs_PCs - COALESCE(prod.IssuedUnits, 0) + COALESCE(ret.ReturnedUnits, 0)) * s.UnitCost AS RemainingInventory

FROM `VuelaProduction.Store` AS s
LEFT JOIN (
    SELECT 
        Inventory,
        SUM(Issued_Kgs_Mtrs_Pcs_) AS IssuedUnits
    FROM `VuelaProduction.Production`
    WHERE Status != 'Cancelled'
    GROUP BY Inventory
) AS prod
    ON s.ID = prod.Inventory
LEFT JOIN (
    SELECT 
        InventoryID,
        SUM(Qty) AS ReturnedUnits
    FROM `VuelaProduction.StoreReturns`
    WHERE Status = 'Approved'
    GROUP BY InventoryID
) AS ret
    ON s.ID = ret.InventoryID
WHERE s.Status = 'Approved';





