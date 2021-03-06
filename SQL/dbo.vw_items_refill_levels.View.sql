USE [sl_inventory]
GO
/****** Object:  View [dbo].[vw_items_refill_levels]    Script Date: 11/9/2018 5:25:41 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

--select * from vw_items_refill_levels
CREATE view [dbo].[vw_items_refill_levels]
--get refill's alerts and warning for all inventory items
as

with received --(order_id, productQty) 
as (
	--list of all received items
	select productCode, sum(productQty) productQty
	from inv_item_received  
	group by productCode
)
, items_in 
as (
	--join of received items to inv_items table
	select --isnull(r.productQty, 0) productQty, isnull (p.unitsPerProduct,0) unitsPerProduct, 
	(isnull(r.productQty, 0) * isnull (p.unitsPerProduct,0)) + isnull(iv.initial_amount,0) itemsIn, --calculates number of available items + initial values
	i.item_id, i.itemName, i.item_units, i.minStockQty, i.refillFactorOfMinStockQty, i.notifyFactorOfMinStockQty
	from inv_items i 
	left join inv_products p on p.item_id = i.item_id
	left join received r on r.productCode = p.productCode
	left join inv_items_initial_values iv on i.item_id = iv.item_id
)
, items_out 
as (
	--create a group by of used items per item_id
	select item_id, sum (itemQty) as itemQty
	from inv_usage
	where isnull(canceled,0) <> 1 and takenDate is not null 
	group by item_id
)
, available 
as (
	--join In table to Out table
	select i.item_id, i.ItemName, 
		i.itemsIn - isnull(o.itemQty,0) itemsAvail,
		i.minStockQty, i.refillFactorOfMinStockQty, i.notifyFactorOfMinStockQty	
	from items_in i 
		left join items_out o on i.item_id = o.item_id
)
--output list of inventory items with the Refill alert and warning flags
select item_id, ItemName, 
itemsAvail, minStockQty, 
iif(minStockQty * refillFactorOfMinStockQty >= itemsAvail, 'Refill', 'OK') as ReFillAlert,
iif(minStockQty * notifyFactorOfMinStockQty >= itemsAvail, 'Refill', 'OK') as ReFillNotify,
cast (minStockQty * refillFactorOfMinStockQty as decimal (8,2)) as refillLevelStockQty, 
cast (minStockQty * notifyFactorOfMinStockQty as decimal (8,2)) as notifyLevelStockQty,
refillFactorOfMinStockQty, notifyFactorOfMinStockQty
from available
GO
