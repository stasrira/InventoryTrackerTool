USE [sl_inventory]
GO
/****** Object:  View [dbo].[vw_items_availability_full]    Script Date: 11/9/2018 5:25:41 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


--select * from vw_items_availability_full
CREATE view [dbo].[vw_items_availability_full]
as 
--availability of items
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
	i.item_id, i.itemName, i.item_units, i.minSampleGroupQty, i.minItemQtyPerSampleGroup
	from inv_items i 
	left join inv_products p on p.item_id = i.item_id
	left join received r on r.productCode = p.productCode
	left join inv_items_initial_values iv on i.item_id = iv.item_id
)
, items_already_taken 
as (
	--create a group by of used items per item_id
	select item_id, sum (itemQty) as itemQty
	from inv_usage
	where isnull(canceled,0) <> 1 and takenDate is not null 
	group by item_id
)
, items_reserved_or_taken 
as (
	--create a group by of used items per item_id
	select item_id, sum (itemQty) as itemQty
	from inv_usage
	where isnull(canceled,0) <> 1 
	group by item_id
)
, items_reserved 
as (
	--create a group by of used items per item_id
	select item_id, sum (itemQty) as itemQty
	from inv_usage
	where isnull(canceled,0) <> 1 and takenDate is null and reservedDate is not null
	group by item_id
)
--join to In table to Out table
select i.item_id, i.ItemName, 
	--i.itemsIn, isnull(u.itemQty,0) itemsOut, 
	i.itemsIn - isnull(t.itemQty,0) itemsAvail_WithoutReserved,
	isnull(r.itemQty,0) itemsReserved,
	i.itemsIn - isnull(rt.itemQty,0) itemsAvail_CountingReserved,
	i.minSampleGroupQty, i.minItemQtyPerSampleGroup
from items_in i 
	left join items_already_taken t on i.item_id = t.item_id
	left join items_reserved_or_taken rt on i.item_id = rt.item_id
	left join items_reserved r on i.item_id = r.item_id

GO
