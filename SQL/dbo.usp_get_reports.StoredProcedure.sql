USE [sl_inventory]
GO
/****** Object:  StoredProcedure [dbo].[usp_get_reports]    Script Date: 11/9/2018 5:25:41 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE procedure [dbo].[usp_get_reports]
as
select report_id, report_name report_name, report_get_SQL, 
isnull (report_cond_format_columns, '') report_cond_format_columns,
isnull (report_cond_format_rules, '') report_cond_format_rules,
isnull (report_action_columns, '') report_action_columns,
isnull (report_action_functions, '') report_action_functions
from inv_reports
where isnull (not_active, 0) = 0
GO
