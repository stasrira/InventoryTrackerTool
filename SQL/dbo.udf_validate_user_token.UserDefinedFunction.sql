USE [sl_inventory]
GO
/****** Object:  UserDefinedFunction [dbo].[udf_validate_user_token]    Script Date: 11/9/2018 5:25:40 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE function [dbo].[udf_validate_user_token] (@token varchar(30)) 
returns int 
as
Begin 
	declare @out int;
	set @out = 1;
	Return @out;

end
GO
