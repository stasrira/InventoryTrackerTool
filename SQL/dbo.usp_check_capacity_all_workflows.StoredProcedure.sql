USE [sl_inventory]
GO
/****** Object:  StoredProcedure [dbo].[usp_check_capacity_all_workflows]    Script Date: 11/9/2018 5:25:41 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

--exec usp_check_capacity_all_workflows 96
CREATE procedure [dbo].[usp_check_capacity_all_workflows] (@sampleQty decimal (8,2))
as 

;with items_capacity 
	as (
	--output capacity estimate per given sample Qty
	Select 
	a.item_id,  
	--a.minSampleGroupQty, a.minItemQtyPerSampleGroup,
	--CEILING(@sampleQty/a.minSampleGroupQty)*a.minItemQtyPerSampleGroup requiredItemsPerSampleEstimated,
	--FLOOR(a.itemsAvail/(CEILING(@sampleQty/a.minSampleGroupQty)*a.minItemQtyPerSampleGroup)) itemCapacityPerSampleEstimated
	FLOOR(a.itemsAvail/dbo.udf_getRequiredItemsPerSampleEstimated (@sampleQty,a.minSampleGroupQty,a.minItemQtyPerSampleGroup)) itemCapacityPerSampleEstimated
	
	from vw_items_availability a 
	)

, workflow_capacity 
	as (
	--output capacity estimate for all workflows
	select workflowID, min (i.itemCapacityPerSampleEstimated) workflowCapacityPerSampleEstimated 
	from items_capacity i 
	inner join inv_item_workflow_sets w on i.item_id = w.item_id
	Group by workflowID
	)

--output workflow capacity with workflow name
Select c.workflowID, w.WorkflowName, 
@sampleQty as [Sample Qty Estimated], 
c.workflowCapacityPerSampleEstimated [Available Capacity]
from workflow_capacity c 
inner join inv_workflows w on c.workflowID = w.workflowID

GO
