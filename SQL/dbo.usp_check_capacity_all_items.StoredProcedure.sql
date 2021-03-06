USE [sl_inventory]
GO
/****** Object:  StoredProcedure [dbo].[usp_check_capacity_all_items]    Script Date: 11/9/2018 5:25:41 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
--exec usp_check_capacity_all_items 96
CREATE procedure [dbo].[usp_check_capacity_all_items] (@sampleQty decimal (8,2))
as 

--output capacity estimate per given sample Qty
Select 
a.item_id, a.ItemName, a.itemsAvail, @sampleQty as [Sample Qty Estimated]
--,a.minSampleGroupQty, a.minItemQtyPerSampleGroup
--, CEILING(@sampleQty/a.minSampleGroupQty)*a.minItemQtyPerSampleGroup requiredItemsPerSampleEstimated_OLD
, dbo.udf_getRequiredItemsPerSampleEstimated (@sampleQty,a.minSampleGroupQty,a.minItemQtyPerSampleGroup) [Items Required Per Estimate]
--, FLOOR(a.itemsAvail/(CEILING(@sampleQty/a.minSampleGroupQty)*a.minItemQtyPerSampleGroup)) itemCapacityPerSampleEstimated_OLD
, FLOOR(a.itemsAvail/dbo.udf_getRequiredItemsPerSampleEstimated (@sampleQty,a.minSampleGroupQty,a.minItemQtyPerSampleGroup)) [Available Capacity]
from vw_items_availability a

GO
