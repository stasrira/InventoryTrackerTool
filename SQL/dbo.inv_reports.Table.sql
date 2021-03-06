USE [sl_inventory]
GO
/****** Object:  Table [dbo].[inv_reports]    Script Date: 11/9/2018 5:25:41 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[inv_reports](
	[report_id] [int] NOT NULL,
	[report_name] [varchar](50) NOT NULL,
	[report_get_SQL] [varchar](200) NOT NULL,
	[report_cond_format_columns] [varchar](200) NULL,
	[report_cond_format_rules] [varchar](200) NULL,
	[report_action_columns] [varchar](200) NULL,
	[report_action_functions] [varchar](200) NULL,
	[not_active] [bit] NULL,
 CONSTRAINT [PK_inv_reports] PRIMARY KEY CLUSTERED 
(
	[report_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
