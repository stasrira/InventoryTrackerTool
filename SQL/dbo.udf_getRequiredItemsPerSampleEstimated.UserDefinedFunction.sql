USE [sl_inventory]
GO
/****** Object:  UserDefinedFunction [dbo].[udf_getRequiredItemsPerSampleEstimated]    Script Date: 11/9/2018 5:25:40 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE function [dbo].[udf_getRequiredItemsPerSampleEstimated] (
@sampleQty int --number of sample being estimated
, @minSampleGroupQty decimal (10,4) --min samples being served by this item
, @minItemQtyPerSampleGroup decimal (10,4) --min amount of item to be used to serve the minSamlpeGroupQty
) 
returns decimal (10,4)
as
Begin 
	declare @out decimal (10,4);
	set @out = CEILING(@sampleQty/@minSampleGroupQty)*@minItemQtyPerSampleGroup;
	Return @out;

end
GO
