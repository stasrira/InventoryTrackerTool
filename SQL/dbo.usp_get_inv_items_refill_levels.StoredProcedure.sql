USE [sl_inventory]
GO
/****** Object:  StoredProcedure [dbo].[usp_get_inv_items_refill_levels]    Script Date: 11/9/2018 5:25:41 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

--exec usp_get_inv_items_refill_levels
CREATE proc [dbo].[usp_get_inv_items_refill_levels] 
as
select * from vw_items_refill_levels
order by ItemName
GO
